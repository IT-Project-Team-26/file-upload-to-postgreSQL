import pandas as pd
import streamlit as st
from sqlalchemy import create_engine, text
from sqlalchemy.exc import SQLAlchemyError
import traceback
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="Data Upload Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better UI
st.markdown("""
    <style>
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 10px;
        margin-bottom: 2rem;
        color: white;
        text-align: center;
    }
    .section-card {
        background-color: white;
        padding: 1.5rem;
        border-radius: 10px;
        margin-bottom: 2rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        border-left: 4px solid #667eea;
    }
    .success-box {
        background-color: #d4edda;
        border-left: 4px solid #28a745;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .error-box {
        background-color: #f8d7da;
        border-left: 4px solid #dc3545;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #d1ecf1;
        border-left: 4px solid #17a2b8;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .stButton button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        font-weight: bold;
        border: none;
        padding: 0.5rem 2rem;
        border-radius: 5px;
        transition: all 0.3s ease;
        width: 100%;
    }
    .stButton button:hover {
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(0,0,0,0.2);
    }
    .upload-header {
        font-size: 1.2rem;
        font-weight: bold;
        margin-bottom: 1rem;
        color: #667eea;
    }
    hr {
        margin: 2rem 0;
    }
    </style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
    <div class="main-header">
        <h1>📊 Database Upload Dashboard</h1>
        <p>Upload all your Excel/CSV files to PostgreSQL in one page</p>
    </div>
""", unsafe_allow_html=True)

# Database configuration in sidebar
with st.sidebar:
    st.markdown("## 🔧 Database Configuration")
    st.markdown("---")
    
    # Database connection parameters
    db_user = st.text_input("👤 Username", value="postgres", help="Database username")
    db_password = st.text_input("🔑 Password", value="123456", type="password", help="Database password")
    db_host = st.text_input("🏠 Host", value="localhost", help="Database host address")
    db_port = st.text_input("🔌 Port", value="5432", help="Database port number")
    db_name = st.text_input("💾 Database Name", value="IT", help="Database name")
    
    st.markdown("---")
    
    # Installation instructions
    with st.expander("📦 Required Packages", expanded=False):
        st.code("""
pip install streamlit pandas sqlalchemy psycopg2-binary openpyxl xlrd
        """)
    
    # Test connection button
    if st.button("🔌 Test Database Connection", use_container_width=True):
        try:
            # Try different connection string formats
            connection_uris = [
                f"postgresql://{db_user}:{db_password}@{db_host}:{db_port}/{db_name}",
                f"postgresql+psycopg2://{db_user}:{db_password}@{db_host}:{db_port}/{db_name}",
                f"postgresql+pg8000://{db_user}:{db_password}@{db_host}:{db_port}/{db_name}"
            ]
            
            connected = False
            for uri in connection_uris:
                try:
                    engine = create_engine(uri)
                    with engine.connect() as conn:
                        conn.execute(text("SELECT 1"))
                    st.success(f"✅ Successfully connected to database using: {uri.split('://')[0]}!")
                    connected = True
                    break
                except:
                    continue
            
            if not connected:
                st.error("❌ Could not connect with any available driver. Please install psycopg2-binary: pip install psycopg2-binary")
                
        except Exception as e:
            st.error(f"❌ Connection failed: {str(e)}")
            st.info("💡 Try installing required package: pip install psycopg2-binary")
    
    st.markdown("---")
    st.markdown("### 📋 Instructions")
    st.markdown("""
    1. Configure database connection above
    2. Upload files for each section below
    3. Click upload button for each section
    4. Check success messages
    """)

# Main content - Single page with all uploaders
connection_params = {
    "user": db_user,
    "password": db_password,
    "host": db_host,
    "port": db_port,
    "database": db_name
}

# Function to read files with multiple encodings and formats
def read_file_with_fallback(file, file_type="excel"):
    """Try to read file with multiple methods and encodings"""
    try:
        if file_type == "excel":
            # Try different engines
            engines = ["openpyxl", "xlrd"]
            for engine in engines:
                try:
                    return pd.read_excel(file, engine=engine)
                except:
                    continue
            raise Exception("Could not read Excel file with any engine")
        else:  # CSV
            encodings = ["utf-8", "latin1", "iso-8859-1", "cp1252"]
            for encoding in encodings:
                try:
                    return pd.read_csv(file, encoding=encoding)
                except:
                    continue
            raise Exception("Could not read CSV file with any encoding")
    except Exception as e:
        st.error(f"Error reading file: {str(e)}")
        return None

# Function to upload data to database
def upload_to_database(df, table_name, connection_params):
    """Upload DataFrame to PostgreSQL with error handling"""
    try:
        # Try different connection formats
        connection_uris = [
            f"postgresql://{connection_params['user']}:{connection_params['password']}@{connection_params['host']}:{connection_params['port']}/{connection_params['database']}",
            f"postgresql+psycopg2://{connection_params['user']}:{connection_params['password']}@{connection_params['host']}:{connection_params['port']}/{connection_params['database']}"
        ]
        
        engine = None
        for uri in connection_uris:
            try:
                engine = create_engine(uri)
                # Test connection
                with engine.connect() as conn:
                    conn.execute(text("SELECT 1"))
                break
            except:
                continue
        
        if engine is None:
            raise Exception("Could not create database connection. Please install psycopg2-binary: pip install psycopg2-binary")
        
        # Check if table exists
        with engine.connect() as conn:
            result = conn.execute(text(f"""
                SELECT EXISTS (
                    SELECT FROM information_schema.tables 
                    WHERE table_schema = 'public' 
                    AND table_name = '{table_name}'
                )
            """))
            table_exists = result.scalar()
        
        # Upload data
        df.to_sql(
            name=table_name,
            con=engine,
            schema="public",
            if_exists="append",
            index=False,
            method='multi',
            chunksize=1000
        )
        
        return {
            "success": True,
            "rows_inserted": len(df),
            "table_exists": table_exists,
            "message": f"Successfully uploaded {len(df)} rows to {table_name}"
        }
    except SQLAlchemyError as e:
        return {
            "success": False,
            "error": str(e),
            "traceback": traceback.format_exc()
        }
    except Exception as e:
        return {
            "success": False,
            "error": str(e),
            "traceback": traceback.format_exc()
        }

# Function to process file uploads
def process_file_upload(uploaded_files, table_name, section_name, connection_params):
    """Process and upload files for a specific table"""
    if not uploaded_files:
        return None
    
    with st.container():
        col1, col2 = st.columns([3, 1])
        
        with col1:
            st.info(f"📁 {len(uploaded_files)} file(s): {', '.join([f.name for f in uploaded_files[:3]])}" + 
                   (f" and {len(uploaded_files)-3} more..." if len(uploaded_files) > 3 else ""))
        
        with col2:
            upload_button = st.button(f"🚀 Upload {section_name}", key=f"btn_{table_name}", type="primary")
        
        if upload_button:
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                dataframes = []
                for i, file in enumerate(uploaded_files):
                    status_text.text(f"Reading file {i+1}/{len(uploaded_files)}: {file.name}")
                    
                    # Check file extension
                    if file.name.endswith('.csv'):
                        df = read_file_with_fallback(file, "csv")
                    else:
                        df = read_file_with_fallback(file, "excel")
                    
                    if df is not None:
                        dataframes.append(df)
                    progress_bar.progress((i + 1) / len(uploaded_files))
                
                if dataframes:
                    status_text.text("Combining dataframes...")
                    combined_df = pd.concat(dataframes, ignore_index=True)
                    
                    # Data cleaning for specific columns
                    if table_name == "binunloading":
                        if "Unloading Started At" in combined_df.columns:
                            combined_df["Unloading Started At"] = pd.to_datetime(
                                combined_df["Unloading Started At"], errors="coerce"
                            )
                        if "Unloading End At" in combined_df.columns:
                            combined_df["Unloading End At"] = pd.to_datetime(
                                combined_df["Unloading End At"], errors="coerce"
                            )
                        if "Duration" in combined_df.columns:
                            combined_df["Duration"] = pd.to_numeric(
                                combined_df["Duration"], errors="coerce"
                            )
                    
                    # Show data preview in expander
                    with st.expander(f"📊 Preview {section_name} Data"):
                        st.dataframe(combined_df.head(10), use_container_width=True)
                        
                        col_a, col_b, col_c = st.columns(3)
                        with col_a:
                            st.metric("Total Rows", len(combined_df))
                        with col_b:
                            st.metric("Total Columns", len(combined_df.columns))
                        with col_c:
                            st.metric("File Count", len(uploaded_files))
                    
                    # Upload to database
                    status_text.text("Uploading to database...")
                    result = upload_to_database(combined_df, table_name, connection_params)
                    progress_bar.progress(100)
                    
                    if result["success"]:
                        st.markdown(f"""
                            <div class="success-box">
                                ✅ {result['message']}<br>
                                📋 Table {'existed' if result['table_exists'] else 'was created'} in database
                            </div>
                        """, unsafe_allow_html=True)
                        status_text.text("Upload completed successfully!")
                        return combined_df
                    else:
                        st.markdown(f"""
                            <div class="error-box">
                                ❌ Upload failed: {result['error']}<br>
                                <details>
                                    <summary>Technical details</summary>
                                    <pre>{result['traceback']}</pre>
                                </details>
                            </div>
                        """, unsafe_allow_html=True)
                        status_text.text("Upload failed!")
                        return None
                else:
                    st.error("No valid data found in uploaded files")
                    return None
                    
            except Exception as e:
                st.markdown(f"""
                    <div class="error-box">
                        ❌ Unexpected error: {str(e)}<br>
                        <details>
                            <summary>Full traceback</summary>
                            <pre>{traceback.format_exc()}</pre>
                        </details>
                    </div>
                """, unsafe_allow_html=True)
                return None
    
    return None

# Check if required packages are installed
try:
    import psycopg2
    st.sidebar.success("✅ psycopg2 is installed")
except ImportError:
    st.sidebar.error("❌ psycopg2 is NOT installed")
    st.sidebar.info("Run this command to install: pip install psycopg2-binary")

# Section 1: Bin Unloading
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="upload-header">📦 1. Bin Unloading Data</div>', unsafe_allow_html=True)
uploaded_binunloading = st.file_uploader(
    "Upload Bin Unloading files (Excel/CSV)",
    type=["xlsx", "xls", "xlsm", "csv"],
    accept_multiple_files=True,
    key="binunloading_uploader"
)
if uploaded_binunloading:
    process_file_upload(uploaded_binunloading, "binunloading", "Bin Unloading", connection_params)
st.markdown('</div>', unsafe_allow_html=True)

# Section 2: KPI 72
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="upload-header">📈 2. KPI 72 Data</div>', unsafe_allow_html=True)
uploaded_kpi72 = st.file_uploader(
    "Upload KPI 72 files (Excel/CSV)",
    type=["xlsx", "xls", "xlsm", "csv"],
    accept_multiple_files=True,
    key="kpi72_uploader"
)
if uploaded_kpi72:
    process_file_upload(uploaded_kpi72, "kpi72", "KPI 72", connection_params)
st.markdown('</div>', unsafe_allow_html=True)

# Section 3: Vehicle Status
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="upload-header">🚗 3. Vehicle Status Report</div>', unsafe_allow_html=True)
uploaded_vehiclestatus = st.file_uploader(
    "Upload Vehicle Status files (Excel/CSV)",
    type=["xlsx", "xls", "xlsm", "csv"],
    accept_multiple_files=True,
    key="vehiclestatus_uploader"
)
if uploaded_vehiclestatus:
    process_file_upload(uploaded_vehiclestatus, "vehicle_status_report", "Vehicle Status", connection_params)
st.markdown('</div>', unsafe_allow_html=True)

# Section 4: GPS Remarks
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="upload-header">📍 4. GPS Remarks</div>', unsafe_allow_html=True)
uploaded_gpsremark = st.file_uploader(
    "Upload GPS Remark files (Excel/CSV)",
    type=["xlsx", "xls", "xlsm", "csv"],
    accept_multiple_files=True,
    key="gpsremark_uploader"
)
if uploaded_gpsremark:
    process_file_upload(uploaded_gpsremark, "gps_remarks", "GPS Remarks", connection_params)
st.markdown('</div>', unsafe_allow_html=True)

# Section 5: KPI 58
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="upload-header">🎯 5. KPI 58 Data</div>', unsafe_allow_html=True)
upload_kpi58 = st.file_uploader(
    "Upload KPI 58 files (Excel/CSV)",
    type=["xlsx", "xls", "xlsm", "csv"],
    accept_multiple_files=True,
    key="kpi58_uploader"
)
if upload_kpi58:
    process_file_upload(upload_kpi58, "kpi58", "KPI 58", connection_params)
st.markdown('</div>', unsafe_allow_html=True)

# Section 6: KPI 52
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="upload-header">⚡ 6. KPI 52 Data</div>', unsafe_allow_html=True)
upload_kpi52 = st.file_uploader(
    "Upload KPI 52 files (Excel/CSV)",
    type=["xlsx", "xls", "xlsm", "csv"],
    accept_multiple_files=True,
    key="kpi52_uploader"
)
if upload_kpi52:
    process_file_upload(upload_kpi52, "kpi52", "KPI 52", connection_params)
st.markdown('</div>', unsafe_allow_html=True)

# Section 7: KPI 52 Current Date
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="upload-header">📅 7. KPI 52 Current Date</div>', unsafe_allow_html=True)
upload_kpi52_currentdate = st.file_uploader(
    "Upload KPI 52 Current Date files (Excel/CSV)",
    type=["xlsx", "xls", "xlsm", "csv"],
    accept_multiple_files=True,
    key="kpi52_currentdate_uploader"
)
if upload_kpi52_currentdate:
    process_file_upload(upload_kpi52_currentdate, "kpi52_current_date", "KPI 52 Current Date", connection_params)
st.markdown('</div>', unsafe_allow_html=True)

# Section 8: KPI 56
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="upload-header">📊 8. KPI 56 Data</div>', unsafe_allow_html=True)
upload_kpi56 = st.file_uploader(
    "Upload KPI 56 files (Excel/CSV)",
    type=["xlsx", "xls", "xlsm", "csv"],
    accept_multiple_files=True,
    key="kpi56_uploader"
)
if upload_kpi56:
    process_file_upload(upload_kpi56, "kpi56", "KPI 56", connection_params)
st.markdown('</div>', unsafe_allow_html=True)

# Section 9: KPI 63a
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="upload-header">🔍 9. KPI 63a Data</div>', unsafe_allow_html=True)
upload_kpi63a = st.file_uploader(
    "Upload KPI 63a files (Excel/CSV)",
    type=["xlsx", "xls", "xlsm", "csv"],
    accept_multiple_files=True,
    key="kpi63a_uploader"
)
if upload_kpi63a:
    process_file_upload(upload_kpi63a, "kpi63a", "KPI 63a", connection_params)
st.markdown('</div>', unsafe_allow_html=True)

# Section 10: KPI 63b
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="upload-header">🔎 10. KPI 63b Data</div>', unsafe_allow_html=True)
upload_kpi63b = st.file_uploader(
    "Upload KPI 63b files (Excel/CSV)",
    type=["xlsx", "xls", "xlsm", "csv"],
    accept_multiple_files=True,
    key="kpi63b_uploader"
)
if upload_kpi63b:
    process_file_upload(upload_kpi63b, "kpi63b", "KPI 63b", connection_params)
st.markdown('</div>', unsafe_allow_html=True)

# Footer with summary
st.markdown("---")
st.markdown("""
    <div style="text-align: center; color: #666; padding: 1rem;">
        <p>📊 Data Upload Dashboard | All sections in one page | Version 2.0</p>
        <p style="font-size: 0.8rem;">💡 Tip: First install required packages: pip install psycopg2-binary</p>
    </div>
""", unsafe_allow_html=True)