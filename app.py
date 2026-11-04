import streamlit as st
from streamlit_sortables import sort_items

st.write("測試拖曳元件")
items = sort_items(["A", "B", "C"])
st.write(items)

multi = sort_items(
    [
        {"header": "可用", "items": ["甲", "乙"]},
        {"header": "已選", "items": ["丙"]},
    ],
    multi_containers=True,
    direction="horizontal",
)
st.write(multi)
