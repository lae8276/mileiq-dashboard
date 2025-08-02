st.subheader("📅 Daily Trip Summary (Double-click any field to copy)")

st.markdown("""
<style>
.copy-table td {
    padding: 6px 12px;
    border: 1px solid #ddd;
    font-family: monospace;
    cursor: pointer;
}
.copy-table tr:nth-child(even) { background-color: #f9f9f9; }
.copy-table tr:hover { background-color: #e0f0ff; }
</style>

<script>
document.addEventListener('DOMContentLoaded', function() {
    document.querySelectorAll('.copy-table td').forEach(cell => {
        cell.ondblclick = () => {
            navigator.clipboard.writeText(cell.innerText.trim());
            cell.style.backgroundColor = '#c4f0c5';
            setTimeout(() => cell.style.backgroundColor = '', 600);
        };
    });
});
</script>
""", unsafe_allow_html=True)

# Build HTML table manually for copy
html = '<table class="copy-table"><thead><tr><th>Date</th><th>Miles</th><th>Postcodes</th></tr></thead><tbody>'
for _, row in summary_df.iterrows():
    html += f"<tr><td>{row['Date']}</td><td>{row['Miles']}</td><td>{row['Postcodes']}</td></tr>"
html += "</tbody></table>"

st.markdown(html, unsafe_allow_html=True)
