# --- STEP 3: FINAL GENERATION ---
st.write("---")
st.subheader("📂 STEP 3: GENERATE DOCUMENT")

# Use a container to keep things organized
if st.button("💾 CLICK TO GENERATE & DOWNLOAD"):
    # Clean the inputs (remove accidental spaces)
    clean_name = name.strip()
    clean_passport = passport.strip()

    if not clean_name:
        st.error("❌ Please enter the Full Name.")
    elif not clean_passport:
        st.error("❌ Please enter the Passport Number.")
    elif not template_file:
        st.error("❌ Template file not assigned. Check your Category selection.")
    else:
        try:
            # Re-verify the context has the latest data
            final_context["name"] = clean_name
            final_context["passport"] = clean_passport
            
            doc = DocxTemplate(template_file)
            doc.render(final_context)
            
            bio = io.BytesIO()
            doc.save(bio)
            bio.seek(0)
            
            st.balloons()
            st.success(f"✅ Document for {clean_name} is ready!")
            st.download_button(
                label=f"📥 Download {category} Letter",
                data=bio,
                file_name=f"{category.replace(' ', '_')}_{clean_name.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"⚠️ System Error: {e}")
