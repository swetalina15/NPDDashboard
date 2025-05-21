import streamlit as st
import requests
import msal
import pandas as pd
from collections import defaultdict
import plotly.express as px

st.set_page_config(page_title="📊 Product Status Tracker", layout="wide")
st.title("📊 Product - Bucket - Status View")

# Auth
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

if not access_token:
    st.error("❌ Authentication failed.")
else:
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    grouped_data = defaultdict(lambda: {
        "Buckets": set(),
        "Statuses": set(),
        "Teams": set(),
        "Links": []
    })

    def task_status_label(task):
        percent = task.get("percentComplete", 0)
        if percent == 100:
            return "✅ Completed"
        elif percent > 0:
            return "🔄 In Progress"
        else:
            return "🟡 Not Started"

    for plan_id in plan_ids:
        # Get plan details
        plan_url = f"https://graph.microsoft.com/v1.0/planner/plans/{plan_id}"
        plan_res = requests.get(plan_url, headers=headers)
        if plan_res.status_code != 200:
            continue

        plan_data = plan_res.json()
        plan_name = plan_data.get("title", f"Plan {plan_id}")
        group_id = plan_data.get("owner", None)
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
            title = task.get("title", "")
            bucket_name = bucket_map.get(task["bucketId"], "Unknown")
            task_id = task.get("id", "")

            task_link = (
                f"https://tasks.office.com/{tenant_id}/en-US/Home/Planner/"
                f"#/plantaskboard?groupId={group_id}&planId={plan_id}&taskId={task_id}"
            )

            grouped_data[title]["Buckets"].add(bucket_name)
            grouped_data[title]["Statuses"].add(status)
            grouped_data[title]["Teams"].add(plan_name)
            grouped_data[title]["Links"].append(f"[{bucket_name}]({task_link})")

    # Convert to DataFrame
    records = []
    for title, data in grouped_data.items():
        records.append({
            "Product Name": title,
            "Buckets": ", ".join(sorted(data["Buckets"])),
            "Statuses": ", ".join(sorted(data["Statuses"])),
            "Team": ", ".join(sorted(data["Teams"])),
            "Open Task Links": " | ".join(data["Links"])
        })

    df = pd.DataFrame(records)

    # ---------------- Filters ------------------
    st.markdown("### 🔎 Filter by Product / Bucket / Team")
    col1, col2, col3 = st.columns(3)

    with col1:
        product_filter = st.selectbox("📦 Product Name", ["All"] + sorted(df["Product Name"].unique().tolist()))

    with col2:
        bucket_filter = st.multiselect("🗂️ Buckets", sorted({b for b_list in df["Buckets"].str.split(", ") for b in b_list}))

    with col3:
        team_filter = st.multiselect("👥 Teams", sorted({t for t_list in df["Team"].str.split(", ") for t in t_list}))

    filtered_df = df.copy()

    if product_filter != "All":
        filtered_df = filtered_df[filtered_df["Product Name"] == product_filter]
    if bucket_filter:
        filtered_df = filtered_df[filtered_df["Buckets"].apply(lambda b: any(bucket in b for bucket in bucket_filter))]
    if team_filter:
        filtered_df = filtered_df[filtered_df["Team"].apply(lambda t: any(team in t for team in team_filter))]

    # ---------------- Final Display ------------------
    st.markdown(f"### 🧮 Total Products: `{filtered_df['Product Name'].nunique()}`")
    st.dataframe(filtered_df.reset_index(drop=True), use_container_width=True)

    # ---------------- Chart Section ------------------
    st.markdown("### 📊 Visual Summary")

    # 🥧 Pie Chart: Product distribution by Team
    st.markdown("#### 👥 Product Distribution by Team")
    team_counts = filtered_df["Team"].str.split(", ").explode().value_counts().reset_index()
    team_counts.columns = ["Team", "Count"]
    if not team_counts.empty:
        fig = px.pie(
            team_counts,
            names='Team',
            values='Count',
            title='🧑‍🤝‍🧑 Team-wise Product Share',
            hole=0.4
        )
        st.plotly_chart(fig, use_container_width=True)
