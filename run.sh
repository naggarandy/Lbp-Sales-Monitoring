#!/bin/bash
cd "$(dirname "$0")"
echo "Installing dependencies..."
pip install -q -r requirements.txt
echo ""
echo "Starting LBP Sales Monitor..."
echo "  → Buka di browser: http://localhost:8501"
echo "  → Dari HP (WiFi sama): http://$(hostname -I 2>/dev/null | awk '{print $1}'):8501"
echo ""
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
