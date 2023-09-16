        
        uploaded_file = st.file_uploader(
            'Letakkan file excel NILAI STANDAR [SEKOLAH KERJASAMA]', type='xlsx')

        if uploaded_file is not None:
            df = pd.read_excel(uploaded_file)

            # 89
            r = df.shape[0]-5
            # 90
            s = df.shape[0]-4
            # 91
            t = df.shape[0]-3
            # 92
            u = df.shape[0]-2

            # JUMLAH PESERTA
            peserta = df.iloc[r, 23]

            # rata-rata jumlah benar
            rata_mat = df.iloc[r, 156]
            rata_ind = df.iloc[r, 157]
            rata_eng = df.iloc[r, 158]
            rata_ipa = df.iloc[r, 159]
            rata_ips = df.iloc[r, 160]
            rata_jml = df.iloc[r, 161]

            # rata-rata nilai standar
            rata_Smat = df.iloc[t, 167]
            rata_Sind = df.iloc[t, 168]
            rata_Seng = df.iloc[t, 169]
            rata_Sipa = df.iloc[t, 170]
            rata_Sips = df.iloc[t, 171]
            rata_Sjml = df.iloc[t, 172]

            max_mat = df.iloc[t, 156]
            max_ind = df.iloc[t, 157]
            max_eng = df.iloc[t, 158]
            max_ipa = df.iloc[t, 159]
            max_ips = df.iloc[t, 160]
            max_jml = df.iloc[t, 161]

            # max nilai standar
            max_Smat = df.iloc[r, 167]
            max_Sind = df.iloc[r, 168]
            max_Seng = df.iloc[r, 169]
            max_Sipa = df.iloc[r, 170]
            max_Sips = df.iloc[r, 171]
            max_Sjml = df.iloc[r, 172]

            # min jumlah benar
            min_mat = df.iloc[u, 156]
            min_ind = df.iloc[u, 157]
            min_eng = df.iloc[u, 158]
            min_ipa = df.iloc[u, 159]
            min_ips = df.iloc[u, 160]
            min_jml = df.iloc[u, 161]

            # min nilai standar
            min_Smat = df.iloc[s, 167]
            min_Sind = df.iloc[s, 168]
            min_Seng = df.iloc[s, 169]
            min_Sipa = df.iloc[s, 170]
            min_Sips = df.iloc[s, 171]
            min_Sjml = df.iloc[s, 172]

            data_jml_benar = {'BIDANG STUDI': ['MATEMATIKA (MAT)', 'B. INDONESIA (IND)', 'B. INGGRIS (ENG)', 'IPA', 'IPS', 'JUMLAH (JML)'],
                              'TERENDAH': [min_mat, min_ind, min_eng, min_ipa, min_ips, min_jml],
                              'RATA-RATA': [rata_mat, rata_ind, rata_eng, rata_ipa, rata_ips, rata_jml],
                              'TERTINGGI': [max_mat, max_ind, max_eng, max_ipa, max_ips, max_jml]}

            jml_benar = pd.DataFrame(data_jml_benar)

            data_n_standar = {'BIDANG STUDI': ['MATEMATIKA (MAT)', 'B. INDONESIA (IND)', 'B. INGGRIS (ENG)', 'IPA', 'IPS', 'JUMLAH (JML)'],
                              'TERENDAH': [min_Smat, min_Sind, min_Seng, min_Sipa, min_Sips, min_Sjml],
                              'RATA-RATA': [rata_Smat, rata_Sind, rata_Seng, rata_Sipa, rata_Sips, rata_Sjml],
                              'TERTINGGI': [max_Smat, max_Sind, max_Seng, max_Sipa, max_Sips, max_Sjml]}

            n_standar = pd.DataFrame(data_n_standar)

            data_jml_peserta = {'JUMLAH PESERTA': [peserta]}

            jml_peserta = pd.DataFrame(data_jml_peserta)

            data_jml_soal = {'BIDANG STUDI': ['MAT', 'IND', 'ENG', 'IPA', 'IPS'],
                             'JUMLAH': [JML_SOAL_MAT, JML_SOAL_IND, JML_SOAL_ENG, JML_SOAL_IPA, JML_SOAL_IPS]}

            jml_soal = pd.DataFrame(data_jml_soal)

            df = df[['LOKASI', 'RANK LOK.', 'RANK NAS.', 'NOMOR NF', 'NAMA SISWA', 'NAMA SEKOLAH', 'KELAS',
                    'MAT', 'IND', 'ENG', 'IPA', 'IPS', 'JML', 'S_MAT', 'S_IND', 'S_ENG', 'S_IPA', 'S_IPS', 'S_JML']]

            # sort best 150
            grouped = df.groupby('LOKASI')
            dfs = []  # List kosong untuk menyimpan DataFrame yang akan digabungkan
            for name, group in grouped:
                dfs.append(group)
            best150 = pd.concat(dfs)

            # sort setiap lokasi
            sort728 = df[df['LOKASI'] == 728]
            sort741 = df[df['LOKASI'] == 741]
            sort743 = df[df['LOKASI'] == 743]
            sort822 = df[df['LOKASI'] == 822]
            sort826 = df[df['LOKASI'] == 826]
            sort846 = df[df['LOKASI'] == 846]
                        
            # best150
            best150_all = best150.sort_values(
                by=['RANK NAS.'], ascending=[True])
            del best150_all['LOKASI']
            del best150_all['RANK LOK.']
            best150_all = best150_all.drop(
                best150_all[(best150_all['RANK NAS.'] > 150)].index)

            # 10 besar setiap lokasi
            # 728
            sort728_10 = sort728.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort728_10['LOKASI']
            sort728_10 = sort728_10.drop(
                sort728_10[(sort728_10['RANK LOK.'] > 10)].index)
            # 741
            sort741_10 = sort741.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort741_10['LOKASI']
            sort741_10 = sort741_10.drop(
                sort741_10[(sort741_10['RANK LOK.'] > 10)].index)
            # 743
            sort743_10 = sort743.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort743_10['LOKASI']
            sort743_10 = sort743_10.drop(
                sort743_10[(sort743_10['RANK LOK.'] > 10)].index)
            # 822
            sort822_10 = sort822.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort822_10['LOKASI']
            sort822_10 = sort822_10.drop(
                sort822_10[(sort822_10['RANK LOK.'] > 10)].index)
            # 826
            sort826_10 = sort826.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort826_10['LOKASI']
            sort826_10 = sort826_10.drop(
                sort826_10[(sort826_10['RANK LOK.'] > 10)].index)
            # 846
            sort846_10 = sort846.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort846_10['LOKASI']
            sort846_10 = sort846_10.drop(
                sort846_10[(sort846_10['RANK LOK.'] > 10)].index)
            
            # All 728
            sort728 = sort728.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort728['LOKASI']
            # All 741
            sort741 = sort741.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort741['LOKASI']
            # All 743
            sort743 = sort743.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort743['LOKASI']
            # All 822
            sort822 = sort822.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort822['LOKASI']
            # All 826
            sort826 = sort826.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort826['LOKASI']
            # All 846
            sort846 = sort846.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort846['LOKASI']
            
            # jumlah row
            # 150
            rowBest150_all = best150_all.shape[0]
            rowBest150 = best150.shape[0]
            # 728
            row728_10 = sort728_10.shape[0]
            row728 = sort728.shape[0]
            # 741
            row741_10 = sort741_10.shape[0]
            row741 = sort741.shape[0]
            # 743
            row743_10 = sort743_10.shape[0]
            row743 = sort743.shape[0]
            # 822
            row822_10 = sort822_10.shape[0]
            row822 = sort822.shape[0]
            # 826
            row826_10 = sort826_10.shape[0]
            row826 = sort826.shape[0]
            # 846
            row846_10 = sort846_10.shape[0]
            row846 = sort846.shape[0]
            
            # Create a Pandas Excel writer using XlsxWriter as the engine.
            # Path file hasil penyimpanan
            file_name = f"{kelas}_{penilaian}_{semester}_sekolah_kerjasama.xlsx"
            file_path = tempfile.gettempdir() + '/' + file_name

            # Menyimpan file Excel
            writer = pd.ExcelWriter(file_path, engine='xlsxwriter')

            # Convert the dataframe to an XlsxWriter Excel object cover.
            jml_benar.to_excel(writer, sheet_name='cover',
                               startrow=10,
                               startcol=0,
                               index=False,
                               )

            # Convert the dataframe to an XlsxWriter Excel object cover.
            n_standar.to_excel(writer, sheet_name='cover',
                               startrow=21,
                               startcol=0,
                               index=False,
                               header=False)

            # Convert the dataframe to an XlsxWriter Excel object cover.
            jml_peserta.to_excel(writer, sheet_name='cover',
                                 startrow=21,
                                 startcol=5,
                                 index=False,
                                 header=False)

            # Convert the dataframe to an XlsxWriter Excel object cover.
            jml_soal.to_excel(writer, sheet_name='cover',
                              startrow=13,
                              startcol=5,
                              index=False,
                              header=False)

            # Ranking 150
            best150_all.to_excel(writer, sheet_name='best_150',
                                 startrow=5,
                                 startcol=0,
                                 index=False,
                                 header=False)

            # 728
            # Convert the dataframe to an XlsxWriter Excel object.
            sort728_10.to_excel(writer, sheet_name='728',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort728.to_excel(writer, sheet_name='728',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 741
            # Convert the dataframe to an XlsxWriter Excel object.
            sort741_10.to_excel(writer, sheet_name='741',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort741.to_excel(writer, sheet_name='741',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 743
            # Convert the dataframe to an XlsxWriter Excel object.
            sort743_10.to_excel(writer, sheet_name='743',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort743.to_excel(writer, sheet_name='743',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 822
            # Convert the dataframe to an XlsxWriter Excel object.
            sort822_10.to_excel(writer, sheet_name='822',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort822.to_excel(writer, sheet_name='822',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 826
            # Convert the dataframe to an XlsxWriter Excel object.
            sort826_10.to_excel(writer, sheet_name='826',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort826.to_excel(writer, sheet_name='826',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 846
            # Convert the dataframe to an XlsxWriter Excel object.
            sort846_10.to_excel(writer, sheet_name='846',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort846.to_excel(writer, sheet_name='846',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            
            # Get the xlsxwriter objects from the dataframe writer object.
            workbook = writer.book

            # membuat worksheet baru
            worksheetcover = writer.sheets['cover']
            worksheetbest = writer.sheets['best_150']
            worksheet728 = writer.sheets['728']
            worksheet741 = writer.sheets['741']
            worksheet743 = writer.sheets['743']
            worksheet822 = writer.sheets['822']
            worksheet826 = writer.sheets['826']
            worksheet846 = writer.sheets['846']
            
            # format workbook
            titleCover = workbook.add_format({
                'bold': 1,
                'border': 0,
                'align': 'left',
                'valign': 'vcenter',
                'font_color': '#00058E',
                'font_size': 52,
                'font_name': 'Arial Black'})
            sub_titleCover = workbook.add_format({
                'bold': 0,
                'border': 0,
                'align': 'left',
                'valign': 'vcenter',
                'font_size': 27,
                'font_name': 'Arial Unicode MS'})
            headerCover = workbook.add_format({
                'bold': 1,
                'border': 0,
                'align': 'left',
                'valign': 'vcenter',
                'font_size': 24,
                'font_name': 'Arial Rounded MT Bold'})
            sub_headerCover = workbook.add_format({
                'bold': 0,
                'border': 0,
                'align': 'left',
                'valign': 'vcenter',
                'font_size': 16,
                'font_name': 'Arial'})
            sub_header1Cover = workbook.add_format({
                'bold': 1,
                'border': 0,
                'align': 'center',
                'valign': 'vcenter',
                'font_size': 20,
                'font_name': 'Arial'})
            kelasCover = workbook.add_format({
                'bold': 1,
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
                'font_size': 40,
                'font_name': 'Arial Rounded MT Bold'})
            borderCover = workbook.add_format({
                'bottom': 1,
                'top': 1,
                'left': 1,
                'right': 1})
            centerCover = workbook.add_format({
                'align': 'center',
                'font_size': 12,
                'font_name': 'Arial'})
            center1Cover = workbook.add_format({
                'align': 'center',
                'font_size': 20,
                'font_name': 'Arial'})
            bodyCover = workbook.add_format({
                'bold': 1,
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
                'font_size': 12,
                'font_name': 'Arial',
                'bg_color': 'FFF684'})

            center = workbook.add_format({
                'align': 'center',
                'font_size': 10,
                'font_name': 'Arial'})
            left = workbook.add_format({
                'align': 'left',
                'font_size': 10,
                'font_name': 'Arial'})
            title = workbook.add_format({
                'bold': 1,
                'border': 0,
                'align': 'center',
                'valign': 'vcenter',
                'font_color': '#00058E',
                'font_size': 12,
                'font_name': 'Arial'})
            sub_title = workbook.add_format({
                'bold': 1,
                'border': 0,
                'align': 'center',
                'valign': 'vcenter',
                'font_size': 12,
                'font_name': 'Arial'})
            subTitle = workbook.add_format({
                'bold': 1,
                'border': 0,
                'align': 'center',
                'valign': 'vcenter',
                'font_size': 14,
                'font_name': 'Arial'})
            header = workbook.add_format({
                'bold': 1,
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Arial',
                'bg_color': 'FFF684'})
            body = workbook.add_format({
                'bold': 0,
                'border': 1,
                'align': 'center',
                'font_size': 10,
                'font_name': 'Arial',
                'bg_color': 'FFF684'})
            border = workbook.add_format({
                'bottom': 1,
                'top': 1,
                'left': 1,
                'right': 1})

            # worksheet cover
            worksheetcover.conditional_format(16, 0, 11, 3,
                                              {'type': 'no_errors', 'format': borderCover})

            worksheetcover.insert_image('F1', r'logo nf.jpg')

            worksheetcover.merge_range('A10:A11', 'BIDANG STUDI', bodyCover)
            worksheetcover.merge_range('B10:B11', 'TERENDAH', bodyCover)
            worksheetcover.merge_range('C10:C11', 'RATA-RATA', bodyCover)
            worksheetcover.merge_range('D10:D11', 'TERTINGGI', bodyCover)
            worksheetcover.merge_range('A20:A21', 'BIDANG STUDI', bodyCover)
            worksheetcover.merge_range('B20:B21', 'TERENDAH', bodyCover)
            worksheetcover.merge_range('C20:C21', 'RATA-RATA', bodyCover)
            worksheetcover.merge_range('D20:D21', 'TERTINGGI', bodyCover)
            worksheetcover.write('F13', 'BIDANG STUDI', bodyCover)
            worksheetcover.merge_range('F20:F21', 'JUMLAH', sub_header1Cover)
            worksheetcover.merge_range('F23:F24', 'PESERTA', sub_header1Cover)
            worksheetcover.write('G13', 'JUMLAH', bodyCover)
            worksheetcover.set_column('A:A', 25.71, centerCover)
            worksheetcover.set_column('B:B', 15, centerCover)
            worksheetcover.set_column('C:C', 15, centerCover)
            worksheetcover.set_column('D:D', 15, centerCover)
            worksheetcover.set_column('F:F', 25.71, centerCover)
            worksheetcover.set_column('G:G', 13, centerCover)
            worksheetcover.merge_range('A1:F3', 'DAFTAR NILAI', titleCover)
            worksheetcover.merge_range(
                'A4:F5', fr'{penilaian}', sub_titleCover)
            worksheetcover.merge_range(
                'A6:F7', fr'{semester} TAHUN {tahun}', headerCover)
            worksheetcover.write('A9', 'JUMLAH BENAR', sub_headerCover)
            worksheetcover.write('A19', 'NILAI STANDAR', sub_headerCover)
            worksheetcover.merge_range('F8:G9', fr'{kelas}', kelasCover)
            worksheetcover.merge_range(
                'F11:G12', 'JUMLAH SOAL', sub_header1Cover)

            worksheetcover.conditional_format(26, 0, 21, 3,
                                              {'type': 'no_errors', 'format': borderCover})

            worksheetcover.conditional_format(17, 6, 13, 5,
                                              {'type': 'no_errors', 'format': borderCover})

            worksheetcover.conditional_format(21, 5, 21, 5,
                                              {'type': 'no_errors', 'format': borderCover})

            # workseet best_150
            worksheetbest.insert_image('A1', r'logo resmi nf.jpg')

            worksheetbest.set_column('A:A', 5.43, center)
            worksheetbest.set_column('B:B', 11.43, center)
            worksheetbest.set_column('C:C', 34.29, left)
            worksheetbest.set_column('D:D', 36.71, left)
            worksheetbest.set_column('E:E', 7.57, left)
            worksheetbest.set_column('F:Q', 6.29, center)
            worksheetbest.merge_range(
                'A1:Q1', fr'150 SISWA KELAS {kelas} PERINGKAT TERTINGGI SEKOLAH KERJASAMA', title)
            worksheetbest.merge_range(
                'A2:Q2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheetbest.merge_range('A4:A5', 'RANK', header)
            worksheetbest.merge_range('B4:B5', 'NOMOR NF', header)
            worksheetbest.merge_range('C4:C5', 'NAMA SISWA', header)
            worksheetbest.merge_range('D4:D5', 'SEKOLAH', header)
            worksheetbest.merge_range('E4:E5', 'KELAS', header)
            worksheetbest.merge_range('F4:K4', 'JUMLAH BENAR', header)
            worksheetbest.merge_range('L4:Q4', 'NILAI STANDAR', header)
            worksheetbest.write('F5', 'MAT', body)
            worksheetbest.write('G5', 'IND', body)
            worksheetbest.write('H5', 'ENG', body)
            worksheetbest.write('I5', 'IPA', body)
            worksheetbest.write('J5', 'IPS', body)
            worksheetbest.write('K5', 'JML', body)
            worksheetbest.write('L5', 'MAT', body)
            worksheetbest.write('M5', 'IND', body)
            worksheetbest.write('N5', 'ENG', body)
            worksheetbest.write('O5', 'IPA', body)
            worksheetbest.write('P5', 'IPS', body)
            worksheetbest.write('Q5', 'JML', body)

            worksheetbest.conditional_format(5, 0, rowBest150_all+4, 16,
                                             {'type': 'no_errors', 'format': border})

            # worksheet 728
            worksheet728.insert_image('A1', r'logo resmi nf.jpg')

            worksheet728.set_column('A:A', 7, center)
            worksheet728.set_column('B:B', 6, center)
            worksheet728.set_column('C:C', 18.14, center)
            worksheet728.set_column('D:D', 25, left)
            worksheet728.set_column('E:E', 13.14, left)
            worksheet728.set_column('F:F', 8.57, center)
            worksheet728.set_column('G:R', 5, center)
            worksheet728.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF POLSEK DEPOK', title)
            worksheet728.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet728.write('A5', 'LOKASI', header)
            worksheet728.write('B5', 'TOTAL', header)
            worksheet728.merge_range('A4:B4', 'RANK', header)
            worksheet728.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet728.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet728.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet728.merge_range('F4:F5', 'KELAS', header)
            worksheet728.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet728.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet728.write('G5', 'MAT', body)
            worksheet728.write('H5', 'IND', body)
            worksheet728.write('I5', 'ENG', body)
            worksheet728.write('J5', 'IPA', body)
            worksheet728.write('K5', 'IPS', body)
            worksheet728.write('L5', 'JML', body)
            worksheet728.write('M5', 'MAT', body)
            worksheet728.write('N5', 'IND', body)
            worksheet728.write('O5', 'ENG', body)
            worksheet728.write('P5', 'IPA', body)
            worksheet728.write('Q5', 'IPS', body)
            worksheet728.write('R5', 'JML', body)

            worksheet728.conditional_format(5, 0, row728_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet728.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF POLSEK DEPOK', title)
            worksheet728.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet728.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet728.write('A22', 'LOKASI', header)
            worksheet728.write('B22', 'TOTAL', header)
            worksheet728.merge_range('A21:B21', 'RANK', header)
            worksheet728.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet728.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet728.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet728.merge_range('F21:F22', 'KELAS', header)
            worksheet728.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet728.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet728.write('G22', 'MAT', body)
            worksheet728.write('H22', 'IND', body)
            worksheet728.write('I22', 'ENG', body)
            worksheet728.write('J22', 'IPA', body)
            worksheet728.write('K22', 'IPS', body)
            worksheet728.write('L22', 'JML', body)
            worksheet728.write('M22', 'MAT', body)
            worksheet728.write('N22', 'IND', body)
            worksheet728.write('O22', 'ENG', body)
            worksheet728.write('P22', 'IPA', body)
            worksheet728.write('Q22', 'IPS', body)
            worksheet728.write('R22', 'JML', body)

            worksheet728.conditional_format(22, 0, row728+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 741
            worksheet741.insert_image('A1', r'logo resmi nf.jpg')

            worksheet741.set_column('A:A', 7, center)
            worksheet741.set_column('B:B', 6, center)
            worksheet741.set_column('C:C', 18.14, center)
            worksheet741.set_column('D:D', 25, left)
            worksheet741.set_column('E:E', 13.14, left)
            worksheet741.set_column('F:F', 8.57, center)
            worksheet741.set_column('G:R', 5, center)
            worksheet741.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF DEPOK 1', title)
            worksheet741.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet741.write('A5', 'LOKASI', header)
            worksheet741.write('B5', 'TOTAL', header)
            worksheet741.merge_range('A4:B4', 'RANK', header)
            worksheet741.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet741.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet741.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet741.merge_range('F4:F5', 'KELAS', header)
            worksheet741.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet741.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet741.write('G5', 'MAT', body)
            worksheet741.write('H5', 'IND', body)
            worksheet741.write('I5', 'ENG', body)
            worksheet741.write('J5', 'IPA', body)
            worksheet741.write('K5', 'IPS', body)
            worksheet741.write('L5', 'JML', body)
            worksheet741.write('M5', 'MAT', body)
            worksheet741.write('N5', 'IND', body)
            worksheet741.write('O5', 'ENG', body)
            worksheet741.write('P5', 'IPA', body)
            worksheet741.write('Q5', 'IPS', body)
            worksheet741.write('R5', 'JML', body)

            worksheet741.conditional_format(5, 0, row741_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet741.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF DEPOK 1', title)
            worksheet741.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet741.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet741.write('A22', 'LOKASI', header)
            worksheet741.write('B22', 'TOTAL', header)
            worksheet741.merge_range('A21:B21', 'RANK', header)
            worksheet741.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet741.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet741.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet741.merge_range('F21:F22', 'KELAS', header)
            worksheet741.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet741.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet741.write('G22', 'MAT', body)
            worksheet741.write('H22', 'IND', body)
            worksheet741.write('I22', 'ENG', body)
            worksheet741.write('J22', 'IPA', body)
            worksheet741.write('K22', 'IPS', body)
            worksheet741.write('L22', 'JML', body)
            worksheet741.write('M22', 'MAT', body)
            worksheet741.write('N22', 'IND', body)
            worksheet741.write('O22', 'ENG', body)
            worksheet741.write('P22', 'IPA', body)
            worksheet741.write('Q22', 'IPS', body)
            worksheet741.write('R22', 'JML', body)

            worksheet741.conditional_format(22, 0, row741+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 743
            worksheet743.insert_image('A1', r'logo resmi nf.jpg')

            worksheet743.set_column('A:A', 7, center)
            worksheet743.set_column('B:B', 6, center)
            worksheet743.set_column('C:C', 18.14, center)
            worksheet743.set_column('D:D', 25, left)
            worksheet743.set_column('E:E', 13.14, left)
            worksheet743.set_column('F:F', 8.57, center)
            worksheet743.set_column('G:R', 5, center)
            worksheet743.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PROKLAMASI DEPOK 2', title)
            worksheet743.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet743.write('A5', 'LOKASI', header)
            worksheet743.write('B5', 'TOTAL', header)
            worksheet743.merge_range('A4:B4', 'RANK', header)
            worksheet743.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet743.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet743.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet743.merge_range('F4:F5', 'KELAS', header)
            worksheet743.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet743.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet743.write('G5', 'MAT', body)
            worksheet743.write('H5', 'IND', body)
            worksheet743.write('I5', 'ENG', body)
            worksheet743.write('J5', 'IPA', body)
            worksheet743.write('K5', 'IPS', body)
            worksheet743.write('L5', 'JML', body)
            worksheet743.write('M5', 'MAT', body)
            worksheet743.write('N5', 'IND', body)
            worksheet743.write('O5', 'ENG', body)
            worksheet743.write('P5', 'IPA', body)
            worksheet743.write('Q5', 'IPS', body)
            worksheet743.write('R5', 'JML', body)

            worksheet743.conditional_format(5, 0, row743_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet743.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF PROKLAMASI DEPOK 2', title)
            worksheet743.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet743.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet743.write('A22', 'LOKASI', header)
            worksheet743.write('B22', 'TOTAL', header)
            worksheet743.merge_range('A21:B21', 'RANK', header)
            worksheet743.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet743.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet743.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet743.merge_range('F21:F22', 'KELAS', header)
            worksheet743.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet743.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet743.write('G22', 'MAT', body)
            worksheet743.write('H22', 'IND', body)
            worksheet743.write('I22', 'ENG', body)
            worksheet743.write('J22', 'IPA', body)
            worksheet743.write('K22', 'IPS', body)
            worksheet743.write('L22', 'JML', body)
            worksheet743.write('M22', 'MAT', body)
            worksheet743.write('N22', 'IND', body)
            worksheet743.write('O22', 'ENG', body)
            worksheet743.write('P22', 'IPA', body)
            worksheet743.write('Q22', 'IPS', body)
            worksheet743.write('R22', 'JML', body)

            worksheet743.conditional_format(22, 0, row743+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 822
            worksheet822.insert_image('A1', r'logo resmi nf.jpg')

            worksheet822.set_column('A:A', 7, center)
            worksheet822.set_column('B:B', 6, center)
            worksheet822.set_column('C:C', 18.14, center)
            worksheet822.set_column('D:D', 25, left)
            worksheet822.set_column('E:E', 13.14, left)
            worksheet822.set_column('F:F', 8.57, center)
            worksheet822.set_column('G:R', 5, center)
            worksheet822.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI SMAIT DARUL QURAN', title)
            worksheet822.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet822.write('A5', 'LOKASI', header)
            worksheet822.write('B5', 'TOTAL', header)
            worksheet822.merge_range('A4:B4', 'RANK', header)
            worksheet822.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet822.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet822.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet822.merge_range('F4:F5', 'KELAS', header)
            worksheet822.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet822.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet822.write('G5', 'MAT', body)
            worksheet822.write('H5', 'IND', body)
            worksheet822.write('I5', 'ENG', body)
            worksheet822.write('J5', 'IPA', body)
            worksheet822.write('K5', 'IPS', body)
            worksheet822.write('L5', 'JML', body)
            worksheet822.write('M5', 'MAT', body)
            worksheet822.write('N5', 'IND', body)
            worksheet822.write('O5', 'ENG', body)
            worksheet822.write('P5', 'IPA', body)
            worksheet822.write('Q5', 'IPS', body)
            worksheet822.write('R5', 'JML', body)

            worksheet822.conditional_format(5, 0, row822_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet822.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI SMAIT DARUL QURAN', title)
            worksheet822.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet822.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet822.write('A22', 'LOKASI', header)
            worksheet822.write('B22', 'TOTAL', header)
            worksheet822.merge_range('A21:B21', 'RANK', header)
            worksheet822.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet822.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet822.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet822.merge_range('F21:F22', 'KELAS', header)
            worksheet822.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet822.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet822.write('G22', 'MAT', body)
            worksheet822.write('H22', 'IND', body)
            worksheet822.write('I22', 'ENG', body)
            worksheet822.write('J22', 'IPA', body)
            worksheet822.write('K22', 'IPS', body)
            worksheet822.write('L22', 'JML', body)
            worksheet822.write('M22', 'MAT', body)
            worksheet822.write('N22', 'IND', body)
            worksheet822.write('O22', 'ENG', body)
            worksheet822.write('P22', 'IPA', body)
            worksheet822.write('Q22', 'IPS', body)
            worksheet822.write('R22', 'JML', body)

            worksheet822.conditional_format(22, 0, row822+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 826
            worksheet826.insert_image('A1', r'logo resmi nf.jpg')

            worksheet826.set_column('A:A', 7, center)
            worksheet826.set_column('B:B', 6, center)
            worksheet826.set_column('C:C', 18.14, center)
            worksheet826.set_column('D:D', 25, left)
            worksheet826.set_column('E:E', 13.14, left)
            worksheet826.set_column('F:F', 8.57, center)
            worksheet826.set_column('G:R', 5, center)
            worksheet826.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI SMAN CMBBS', title)
            worksheet826.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet826.write('A5', 'LOKASI', header)
            worksheet826.write('B5', 'TOTAL', header)
            worksheet826.merge_range('A4:B4', 'RANK', header)
            worksheet826.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet826.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet826.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet826.merge_range('F4:F5', 'KELAS', header)
            worksheet826.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet826.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet826.write('G5', 'MAT', body)
            worksheet826.write('H5', 'IND', body)
            worksheet826.write('I5', 'ENG', body)
            worksheet826.write('J5', 'IPA', body)
            worksheet826.write('K5', 'IPS', body)
            worksheet826.write('L5', 'JML', body)
            worksheet826.write('M5', 'MAT', body)
            worksheet826.write('N5', 'IND', body)
            worksheet826.write('O5', 'ENG', body)
            worksheet826.write('P5', 'IPA', body)
            worksheet826.write('Q5', 'IPS', body)
            worksheet826.write('R5', 'JML', body)

            worksheet826.conditional_format(5, 0, row826_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet826.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI SMAN CMBBS', title)
            worksheet826.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet826.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet826.write('A22', 'LOKASI', header)
            worksheet826.write('B22', 'TOTAL', header)
            worksheet826.merge_range('A21:B21', 'RANK', header)
            worksheet826.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet826.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet826.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet826.merge_range('F21:F22', 'KELAS', header)
            worksheet826.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet826.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet826.write('G22', 'MAT', body)
            worksheet826.write('H22', 'IND', body)
            worksheet826.write('I22', 'ENG', body)
            worksheet826.write('J22', 'IPA', body)
            worksheet826.write('K22', 'IPS', body)
            worksheet826.write('L22', 'JML', body)
            worksheet826.write('M22', 'MAT', body)
            worksheet826.write('N22', 'IND', body)
            worksheet826.write('O22', 'ENG', body)
            worksheet826.write('P22', 'IPA', body)
            worksheet826.write('Q22', 'IPS', body)
            worksheet826.write('R22', 'JML', body)

            worksheet826.conditional_format(22, 0, row826+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 846
            worksheet846.insert_image('A1', r'logo resmi nf.jpg')

            worksheet846.set_column('A:A', 7, center)
            worksheet846.set_column('B:B', 6, center)
            worksheet846.set_column('C:C', 18.14, center)
            worksheet846.set_column('D:D', 25, left)
            worksheet846.set_column('E:E', 13.14, left)
            worksheet846.set_column('F:F', 8.57, center)
            worksheet846.set_column('G:R', 5, center)
            worksheet846.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI SMAIT BUAHATI', title)
            worksheet846.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet846.write('A5', 'LOKASI', header)
            worksheet846.write('B5', 'TOTAL', header)
            worksheet846.merge_range('A4:B4', 'RANK', header)
            worksheet846.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet846.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet846.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet846.merge_range('F4:F5', 'KELAS', header)
            worksheet846.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet846.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet846.write('G5', 'MAT', body)
            worksheet846.write('H5', 'IND', body)
            worksheet846.write('I5', 'ENG', body)
            worksheet846.write('J5', 'IPA', body)
            worksheet846.write('K5', 'IPS', body)
            worksheet846.write('L5', 'JML', body)
            worksheet846.write('M5', 'MAT', body)
            worksheet846.write('N5', 'IND', body)
            worksheet846.write('O5', 'ENG', body)
            worksheet846.write('P5', 'IPA', body)
            worksheet846.write('Q5', 'IPS', body)
            worksheet846.write('R5', 'JML', body)

            worksheet846.conditional_format(5, 0, row846_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet846.merge_range(
                'A17:R17', fr'KELAS {kelas} - SMAIT BUAHATI', title)
            worksheet846.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet846.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet846.write('A22', 'LOKASI', header)
            worksheet846.write('B22', 'TOTAL', header)
            worksheet846.merge_range('A21:B21', 'RANK', header)
            worksheet846.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet846.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet846.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet846.merge_range('F21:F22', 'KELAS', header)
            worksheet846.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet846.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet846.write('G22', 'MAT', body)
            worksheet846.write('H22', 'IND', body)
            worksheet846.write('I22', 'ENG', body)
            worksheet846.write('J22', 'IPA', body)
            worksheet846.write('K22', 'IPS', body)
            worksheet846.write('L22', 'JML', body)
            worksheet846.write('M22', 'MAT', body)
            worksheet846.write('N22', 'IND', body)
            worksheet846.write('O22', 'ENG', body)
            worksheet846.write('P22', 'IPA', body)
            worksheet846.write('Q22', 'IPS', body)
            worksheet846.write('R22', 'JML', body)

            worksheet846.conditional_format(22, 0, row846+21, 17,
                                            {'type': 'no_errors', 'format': border})

            workbook.close()
            st.success("File siap diunduh!")

            # Tombol unduh file
            with open(file_path, "rb") as f:
                bytes_data = f.read()
            st.download_button(label="Unduh File", data=bytes_data,
                               file_name=file_name)