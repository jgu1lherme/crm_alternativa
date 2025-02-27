st.markdown(
    """
    <style>
    .stApp {
        background-color: #0e1117;
    }
    h1, h2, h3, p, label {
        color: white;

        /* Personalizando a barra lateral */
    section[data-testid="stSidebar"] {
        background-color: #16161e !important;  
    } 

    }
    .stDataFrame {
        background-color: #1c2025;
        border-radius: 10px;
        padding: 10px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.sidebar.markdown(
    """
    <style>
    /* Força a cor de fundo da sidebar */
    [data-testid="stSidebar"] {
        background-color: #1c1e26 !important;
    }
    """,
    unsafe_allow_html=True,
)