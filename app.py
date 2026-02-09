import streamlit as st
import requests
import msal
import pandas as pd
import plotly.express as px

# ---------------- Page Config ----------------
st.set_page_config(page_title="📊 Product Status Tracker", layout="wide")
st.title("📊 NPD Dashboard")

# ---------------- Auth ----------------
client_id = st.secrets["CLIENT_ID"]
client_secret = st.secrets["CLIENT_SECRET"]
tenant_id = st.secrets["TENANT_ID"]

authority = f"https://login.microsoftonline.com/{tenant_id}"
scope = ["https://graph.microsoft.com/.default"]

app = msal.ConfidentialClientApplication(
    client_id,
    authority=authority,
    client_credential=client_secret,
)

token_response = app.acquire_token_for_client(scopes=scope)
access_token = token_response.get("access_token")

# ---------------- Planner Plan IDs ----------------
plan_ids = [
    "-dg9FJCoHkeg04AlKb_22ckAB08q",
    "1qTmx04ZQ0aUmfMRl-qDAMkAAShd",
    "9MwY0H0E1UipbdU_MQN1pskACY44",
    "HZUriORIbU2o6gb5wRpcPskAAOku",
    "LcvQROmlP0mjBFaizgn-6MkACnHV",
    "PJVx-ra-lU65RVcF_zOPcMkAHDIm",
    "Q-dOJFb1SkiuSMQiCIEZ2ckAEcKR",
    "SjFKBXJCqkucjHDUXmqfFckADR6Y",
    "_CSis4zCf0eODLqCuYG2iskACLvW",
    "hO9_bkDTgES372fKeT0QZckAC9JU",
    "rPvsaKHA3Eqt5QpO1TAlGckAEJEU",
    "s1IswOPOxkWD8AXZOv6EmskABJ4o",
    "Ny5u_Gfh9kygH1HZ4xOGKckABUX7",
]

# ---------------- Auth Check ----------------
if not access_token:
    st.error("❌ Authentication failed.")
    st.stop()

headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json"
}

task_rows = []

# ---------------- Helpers ----------------
def task_status_label(task):
    percent = task.get("percentComplete", 0)
    if percent == 100:
        return "✅ Completed"
    elif percent > 0:
        return "🔄 In Progress"
    else:
        return "🟡 Not Started"

# ---------------- Fetch Planner Data ----------------
for plan_id in plan_ids:

    # Plan Info
    plan_url = f"https://graph.microsoft.com/v1.0/planner/plans/{plan_id}"
    plan_res = requests.get(plan_url, headers=headers)
    if plan_res.status_code != 200:
        continue

    plan_data = plan_res.json()
    plan_name = plan_data.get("title", f"Plan {plan_id}")
    group_id = plan_data.get("owner")
    if not group_id:
        continue

    # Buckets
    bucket_url = f"https://graph.microsoft.com/v1.0/planner/plans/{plan_id}/buckets"
    bucket_res = requests.get(bucket_url, headers=headers)
    if bucket_res.status_code != 200:
        continue

    buckets = bucket_res.json().get("value", [])
    bucket_map = {b["id"]: b["name"] for b in buckets}

    # Tasks
    task_url = f"https://graph.microsoft.com/v1.0/planner/plans/{plan_id}/tasks"
    task_res = requests.get(task_url, headers=headers)
    if task_res.status_code != 200:
        continue

    tasks = task_res.json().get("value", [])

    for task in tasks:

        status = task_status_label(task)
        if "Completed" in status:
            continue

        task_id = task.get("id", "")
        title = task.get("title", "")
        bucket_id = task.get("bucketId", "")
        bucket_name = bucket_map.get(bucket_id, "Unknown")

        created_date = task.get("createdDateTime")
        due_date = task.get("dueDateTime")

        task_link = (
            f"https://tasks.office.com/{tenant_id}/en-US/Home/Planner/"
            f"#/plantaskboard?groupId={group_id}&planId={plan_id}&taskId={task_id}"
        )

        task_rows.append({
            "Product Name": title,
            "Bucket": bucket_name,
            "Status": status,
            "Team": plan_name,
            "Created Date": created_date,
            "Due Date": due_date,
            "Open Task Link": f"[{bucket_name}]({task_link})"
        })

# ---------------- DataFrame ----------------
df = pd.DataFrame(task_rows)

# Safe datetime handling
df["Created Date"] = pd.to_datetime(df["Created Date"], errors="coerce")
df["Due Date"] = pd.to_datetime(df["Due Date"], errors="coerce")

df["Created Date"] = df["Created Date"].dt.strftime("%d-%b-%Y")
df["Due Date"] = df["Due Date"].dt.strftime("%d-%b-%Y")

df["Due Date"] = df["Due Date"].fillna("No Due Date")

# ---------------- Filters ----------------
st.markdown("### 🔎 Filter by Product / Bucket / Team")

col1, col2, col3 = st.columns(3)

with col1:
    product_filter = st.selectbox(
        "📦 Product Name",
        ["All"] + sorted(df["Product Name"].dropna().unique().tolist())
    )

with col2:
    bucket_filter = st.multiselect(
        "🗂️ Buckets",
        sorted(df["Bucket"].dropna().unique().tolist())
    )

with col3:
    team_filter = st.multiselect(
        "👥 Teams",
        sorted(df["Team"].dropna().unique().tolist())
    )

filtered_df = df.copy()

if product_filter != "All":
    filtered_df = filtered_df[filtered_df["Product Name"] == product_filter]

if bucket_filter:
    filtered_df = filtered_df[filtered_df["Bucket"].isin(bucket_filter)]

if team_filter:
    filtered_df = filtered_df[filtered_df["Team"].isin(team_filter)]

# ---------------- Display ----------------
st.markdown(f"### 🧮 Total Products: `{filtered_df['Product Name'].nunique()}`")

st.dataframe(
    filtered_df.reset_index(drop=True),
    use_container_width=True
)

# ---------------- Charts ----------------
st.markdown("### 📊 Visual Summary")

st.markdown("#### 👥 Product Distribution by Team")
team_counts = filtered_df["Team"].value_counts().reset_index()
team_counts.columns = ["Team", "Count"]

if not team_counts.empty:
    fig = px.pie(
        team_counts,
        names="Team",
        values="Count",
        hole=0.4,
        title="🧑‍🤝‍🧑 Team-wise Product Share"
    )
    st.plotly_chart(fig, use_container_width=True)
