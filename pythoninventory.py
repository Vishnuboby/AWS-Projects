import streamlit as st
import pandas as pd
from pyVim.connect import SmartConnect, Disconnect
from pyVmomi import vim
import ssl
import io
from datetime import datetime

# -----------------------------
# Function to collect VM data
# -----------------------------
def get_vm_details(vcenter_ip, username, password, portgroups):
    context = ssl._create_unverified_context()
    si = SmartConnect(host=vcenter_ip, user=username, pwd=password, sslContext=context)
    content = si.RetrieveContent()

    vms = []
    
    # distributed portgroup lookup
    dvs_map = {}
    dv_view = content.viewManager.CreateContainerView(content.rootFolder,[vim.dvs.DistributedVirtualPortgroup],True)
    for pg in dv_view.view:
        dvs_map[pg.key] = pg.name
    dv_view.Destroy()

    # VM view
    vm_view = content.viewManager.CreateContainerView(content.rootFolder,[vim.VirtualMachine],True)

    for vm in vm_view.view:
        if vm.config is None or vm.config.template:
            continue

        summary = vm.summary

        for dev in vm.config.hardware.device:

            # standard
            if hasattr(dev,"backing") and hasattr(dev.backing,"network"):
                if dev.backing.network and dev.backing.network.name in portgroups:
                    vms.append({
                        "VCENTER": vcenter_ip,
                        "VM Name": summary.config.name,
                        "IP": summary.guest.ipAddress,
                        "DNS": summary.guest.hostName,
                        "Memory (GB)": summary.config.memorySizeMB/1024,
                        "CPU": summary.config.numCpu,
                        "Provisioned (GB)": summary.storage.committed/(1024**3),
                        "Used (GB)": summary.storage.uncommitted/(1024**3) if summary.storage.uncommitted else 0,
                        "Guest OS": summary.config.guestFullName,
                        "Power": summary.runtime.powerState
                    })
                    break

            # distributed
            if hasattr(dev,"backing") and hasattr(dev.backing,"port"):
                pg_key = dev.backing.port.portgroupKey
                pg_name = dvs_map.get(pg_key)

                if pg_name in portgroups:
                    vms.append({
                        "VCENTER": vcenter_ip,
                        "VM Name": summary.config.name,
                        "IP": summary.guest.ipAddress,
                        "DNS": summary.guest.hostName,
                        "Memory (GB)": summary.config.memorySizeMB/1024,
                        "CPU": summary.config.numCpu,
                        "Provisioned (GB)": summary.storage.committed/(1024**3),
                        "Used (GB)": summary.storage.uncommitted/(1024**3) if summary.storage.uncommitted else 0,
                        "Guest OS": summary.config.guestFullName,
                        "Power": summary.runtime.powerState
                    })
                    break

    vm_view.Destroy()
    Disconnect(si)

    df = pd.DataFrame(vms).drop_duplicates(subset=["VM Name"])
    return df


# -----------------------------
# WEB UI
# -----------------------------
st.title("VM Inventory Collector")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
if uploaded_file:
    df_excel = pd.read_excel(uploaded_file)
    st.success("Excel file loaded")

    customers = sorted(df_excel["CustomerName"].unique().tolist())
    selected_customers = st.multiselect("Select Customers", customers)

    vcenters = st.text_area("Enter vCenter IPs (one per line)").splitlines()
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Fetch Inventory"):
        if not selected_customers or not vcenters:
            st.error("Please select customers and enter vCenters")
        else:
            output = {}
            with st.spinner("Collecting VM inventory..."):

                for cust in selected_customers:
                    portgroups = df_excel[df_excel["CustomerName"] == cust]["PortGroupName"].tolist()
                    
                    df_final = pd.DataFrame()

                    for vc in vcenters:
                        df_tmp = get_vm_details(vc, username, password, portgroups)
                        df_final = pd.concat([df_final, df_tmp], ignore_index=True)

                    output[cust] = df_final

            st.success("Scan completed!")

            # create excel in memory
            excel_bytes = io.BytesIO()
            with pd.ExcelWriter(excel_bytes, engine='xlsxwriter') as writer:
                for cust, data in output.items():
                    data.to_excel(writer, sheet_name=cust, index=False)
            excel_bytes.seek(0)

            file_name = f"Inventory_{datetime.now().strftime('%Y-%m-%d')}.xlsx"

            st.download_button(
                label="Download Inventory",
                data=excel_bytes,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
