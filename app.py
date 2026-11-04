import streamlit as st
from streamlit_sortables import sort_items

data = lambda: [
    {"header": "可用", "items": ["甲", "乙"]},
    {"header": "已選", "items": ["丙"]},
]

st.write("① 多容器、預設方向(垂直)")
st.write(sort_items(data(), multi_containers=True, key="t1"))

st.write("② 多容器、horizontal")
st.write(sort_items(data(), multi_containers=True, direction="horizontal", key="t2"))

st.write("③ 多容器、horizontal、自訂樣式")
style = """
.sortable-component { display: flex; flex-direction: row; gap: 12px; }
.sortable-container { flex: 1; min-height: 60px; }
"""
st.write(sort_items(data(), multi_containers=True, direction="horizontal",
                    custom_style=style, key="t3"))
