import streamlit as st
import pandas as pd
import simplekml
from io import BytesIO
import zipfile

def create_full_structure(kml):
    """Membuat struktur folder lengkap dengan semua subfolder"""
    # Folder EXISTING dan semua subfoldernya
    existing = kml.newfolder(name="EXISTING")
    existing.newfolder(name="ODP")
    existing.newfolder(name="TIANG")
    existing.newfolder(name="DISTRIBUSI")
    existing.newfolder(name="BOUNDARY")
    existing.newfolder(name="ODC")
    existing.newfolder(name="CLOSURE")
    existing.newfolder(name="FEEDER")
    
    # Folder NEW PLANNING dan semua subfoldernya
    new_planning = kml.newfolder(name="NEW PLANNING")
    new_planning.newfolder(name="ODP")
    new_planning.newfolder(name="TIANG")
    new_planning.newfolder(name="DISTRIBUSI")
    new_planning.newfolder(name="BOUNDARY")
    new_planning.newfolder(name="ODC")
    new_planning.newfolder(name="FEEDER")
    new_planning.newfolder(name="CLOSURE")
    
    # Folder HOUSEHOLD
    kml.newfolder(name="HOUSEHOLD")
    
    return kml

def main():
    st.title("Konversi Excel ke KML (Struktur Lengkap)")
    st.write("Upload file Excel untuk menghasilkan KML dengan struktur folder lengkap")

    # URL ikon
    ODP_ICON = "http://maps.google.com/mapfiles/kml/paddle/ltblu-stars.png"
    HOUSE_ICON = "http://maps.google.com/mapfiles/kml/shapes/homegardenbusiness.png"

    uploaded_file = st.file_uploader("Pilih file Excel", type=["xlsx", "xls"])

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            required_columns = ['NAMA PROJECT', 'Deskripsi', 'ODP', 'LAT ODP', 'LONG ODP', 
                              'name', 'LAT PELANGGAN', 'LONG PELANGGAN']
            
            if not all(col in df.columns for col in required_columns):
                st.error(f"Kolom wajib: {', '.join(required_columns)}")
                return

            zip_buffer = BytesIO()
            
            with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
                progress_bar = st.progress(0)
                projects = df['NAMA PROJECT'].unique()
                total_projects = len(projects)
                
                for i, project_name in enumerate(projects):
                    kml = simplekml.Kml()
                    kml = create_full_structure(kml)  # Buat struktur lengkap
                    
                    # Dapatkan referensi folder yang akan diisi
                    odp_folder = kml.document.folder[0].folder[0]  # EXISTING > ODP
                    household_folder = kml.document.folder[2]       # HOUSEHOLD
                    
                    # Isi data untuk project ini
                    project_data = df[df['NAMA PROJECT'] == project_name]
                    
                    for _, row in project_data.iterrows():
                        # Tambahkan ODP
                        odp = odp_folder.newpoint(
                            name=row['ODP'],
                            description=f"Deskripsi: {row['Deskripsi']}\nProject: {row['NAMA PROJECT']}"
                        )
                        odp.coords = [(row['LONG ODP'], row['LAT ODP'])]
                        odp.style.iconstyle.icon.href = ODP_ICON
                        
                        # Tambahkan Pelanggan
                        house = household_folder.newpoint(name=row['name'])
                        house.coords = [(row['LONG PELANGGAN'], row['LAT PELANGGAN'])]
                        house.style.iconstyle.icon.href = HOUSE_ICON
                    
                    zip_file.writestr(f"{project_name}.kml", kml.kml())
                    progress_bar.progress((i + 1) / total_projects)
            
            st.success(f"Berhasil membuat {total_projects} file KML!")
            
            # Tampilkan contoh struktur
            with st.expander("**Struktur Folder KML**"):
                st.code("""
                EXISTING/
                ├── ODP/         (berisi placemark ODP)
                ├── TIANG/       (kosong)
                ├── DISTRIBUSI/  (kosong)
                ├── BOUNDARY/    (kosong)
                ├── ODC/         (kosong)
                ├── CLOSURE/      (kosong)
                └── FEEDER/       (kosong)
                
                NEW PLANNING/
                ├── ODP/         (kosong)
                ├── TIANG/       (kosong)
                ├── ...         (dst)
                
                HOUSEHOLD/       (berisi placemark pelanggan)
                """)
            
            st.download_button(
                label="Download KML (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="projects_full_structure.zip",
                mime="application/zip"
            )
            
        except Exception as e:
            st.error(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
