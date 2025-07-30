import streamlit as st
import os
from main_script import download_data, process_data, generate_visualization
import time

def main():
    st.title("Maintenance Report Visualization")
    st.write("Enter your credentials below to download the latest data and generate the visualization.")

    username = st.text_input("Username", placeholder="Enter your Propertyware username or email")
    password = st.text_input("Password", type="password", placeholder="Enter your Propertyware password")

    if st.button("Refresh Data and Plot"):
        if username and password:
            st.write("**Overall Progress**")
            overall_progress = st.progress(0)  # Initialize the overall progress bar

            st.write("**Current Step Progress**")
            step_progress = st.progress(0)  # Progress bar for Step 1

            # Step 1: Log into the website and download
            st.write("**Step 1: Log in and Retrieve the Latest Data**")
            with st.spinner("Connecting to the website and downloading the latest data..."):
                try:
                    for progress, message in download_data(username, password):
                        step_progress.progress(progress)
                        st.write(f"- {message}")  # Update with intermediate messages
                    st.success("Step 1 completed: Data retrieved successfully!")
                except Exception as e:
                    st.error(f"Error during login: {e}")
                    return
            overall_progress.progress(33)  # Update overall progress
            step_progress.empty()  # Clear step-specific progress bar
            time.sleep(1)

            # Step 2: Process data
            st.write("**Step 2: Process and Analyze Data**")
            with st.spinner("Processing the downloaded data..."):
                step_progress = st.progress(0)
                for i in range(101):
                    time.sleep(0.05)  # Simulate processing time
                    step_progress.progress(i)
                file_path = os.path.expanduser("~/Downloads/30_Day_Maintenance_Metrics.xlsx")
                if os.path.exists(file_path):
                    data, vendor_summary = process_data(file_path)
                    st.success("Step 2 completed: Data processed successfully.")
                else:
                    st.error("File not found. Ensure the data was downloaded.")
                    return
            overall_progress.progress(66)  # Update overall progress
            step_progress.empty()
            time.sleep(1)

            # Step 3: Generate and display the visualization
            st.write("**Step 3: Generate and Display Visualization**")
            with st.spinner("Creating the graph..."):
                step_progress = st.progress(0)
                for i in range(101):
                    time.sleep(0.03)  # Simulate visualization time
                    step_progress.progress(i)
                generate_visualization(data)
                st.success("Step 3 completed: Visualization generated!")
            # Display Word Frequency Plot
            st.image("top_words_plot.png", caption="Word Frequency Visualization")

            # Display Sentiment Distribution Plot
            st.image("sentiment_distribution.png", caption="Sentiment Distribution Across Work Orders")

            # Display Priority Distribution Plot
            st.image("priority_distribution.png", caption="Work Order Priority Distribution")

            # Display Prioritized Work Orders as a Table
            st.write("### Prioritized Work Orders")
            high_priority = data[['WO#', 'Description', 'Sentiment', 'Priority']].sort_values(by='Sentiment')
            st.dataframe(high_priority)
            overall_progress.progress(100)  # Update overall progress
            step_progress.empty()

            # Display Vendor Summary Table
            st.write("### Vendor Summary")
            st.write("Table showing the count of work orders and average completion duration (in days) by each vendor.")
            if 'vendor_summary' in locals():
                st.dataframe(vendor_summary)
            else:
                st.warning("No vendor summary available.")


            # Cleanup
            try:
                os.remove(file_path)
                st.info("Temporary file cleaned up.")
            except OSError as e:
                st.warning(f"Could not delete file: {e}")
        else:
            st.error("Please enter both username and password.")

if __name__ == "__main__":
    main()
