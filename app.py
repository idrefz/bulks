import streamlit as st
import pandas as pd
import simplekml
from io import BytesIO
import zipfile
import base64

# Fungsi untuk membuat ikon custom dalam format Base64
def get_custom_icon(color="ff0000", icon_type="odp"):
    """Mengembalikan URL ikon custom dalam format base64"""
    if icon_type == "odp":
        # Ikon ODP (misal: kotak merah)
        svg = f"""<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24">
                <rect width="24" height="24" fill="#{color}" stroke="#000"/>
                <text x="12" y="12" font-size="10" fill="white" text-anchor="middle" dominant-baseline="middle">ODP</text>
                </svg>"""
    else:
        # Ikon Household (misal: rumah biru)
        svg = f"""<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24">
                <path d="M12 2L2 7v13h20V7L12 2zm0 2.5L18 7v3h-2V8h-4v2H6V7l6-2.5z" fill="#{color}"/>
                </svg>"""
    
    return f"data:image/svg+xml;base64,{base64.b64encode(svg.encode()).decode()}"

def create_kml_structure(kml):
    """Membuat struktur folder dengan ikon custom"""
    # Ikon custom
    odp_icon = get_custom_icon("FF0000", "odp")  # Merah untuk ODP
    house_icon = get_custom_icon("0000FF", "house")  # Biru untuk Household
    
    # Folder EXISTING > ODP
    existing = kml.newfolder(name="EXISTING")
    odp_folder = existing.newfolder(name="ODP")
    
    # Folder HOUSEHOLD
    household_folder = kml.newfolder(name="HOUSEHOLD")
    
    return odp_folder, household_folder, odp_icon, house_icon

def main():
    st.title("Konversi Excel ke KML dengan Ikon Custom")
    
    uploaded_file = st.file_uploader("Upload file Excel", type=["xlsx", "xls"])
    
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            required_cols = ['NAMA PROJECT', 'ODP', 'LAT ODP', 'LONG ODP', 'name', 'LAT PELANGGAN', 'LONG PELANGGAN']
            
            if not all(col in df.columns for col in required_cols):
                st.error(f"Kolom wajib: {', '.join(required_cols)}")
                return
                
            zip_buffer = BytesIO()
            
            with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
                for project_name, group in df.groupby('NAMA PROJECT'):
                    kml = simplekml.Kml()
                    odp_folder, household_folder, odp_icon, house_icon = create_kml_structure(kml)
                    
                    # Tambahkan ODP dengan ikon custom
                    for _, row in group.iterrows():
                        # Placemark ODP
                        odp = odp_folder.newpoint(name=row['ODP'])
                        odp.coords = [(row['LONG ODP'], row['LAT ODP'])]
                        odp.style.iconstyle.icon.href = odp_icon
                        
                        # Placemark Household
                        house = household_folder.newpoint(name=row['name'])
                        house.coords = [(row['LONG PELANGGAN'], row['LAT PELANGGAN'])]
                        house.style.iconstyle.icon.href = house_icon
                    
                    zip_file.writestr(f"{project_name}.kml", kml.kml())
            
            st.success("Konversi selesai!")
            
            # Tampilkan preview ikon
            col1, col2 = st.columns(2)
            with col1:
                st.image(get_custom_icon("FF0000", "odp"), caption="Ikon ODP", width=50)
            with col2:
                st.image(get_custom_icon("0000FF", "house"), caption="Ikon Household", width=50)
            
            st.download_button(
                "Download KML (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="projects_with_custom_icons.zip",
                mime="application/zip"
            )
            
        except Exception as e:
            st.error(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
