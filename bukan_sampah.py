        
        uploaded_file = st.file_uploader(
            'Letakkan file excel NILAI STANDAR [LOKASI DEPOK-PADANG]', type='xlsx')

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
            sort530 = df[df['LOKASI'] == 530]
            sort531 = df[df['LOKASI'] == 531]
            sort532 = df[df['LOKASI'] == 532]
            sort533 = df[df['LOKASI'] == 533]
            sort534 = df[df['LOKASI'] == 534]
            sort535 = df[df['LOKASI'] == 535]
            sort546 = df[df['LOKASI'] == 546]
            sort547 = df[df['LOKASI'] == 547]
            sort548 = df[df['LOKASI'] == 548]
            sort549 = df[df['LOKASI'] == 549]
            sort556 = df[df['LOKASI'] == 556]
            sort557 = df[df['LOKASI'] == 557]
            sort558 = df[df['LOKASI'] == 558]
            sort575 = df[df['LOKASI'] == 575]
            sort576 = df[df['LOKASI'] == 576]
            sort577 = df[df['LOKASI'] == 577]
            sort578 = df[df['LOKASI'] == 578]
            sort588 = df[df['LOKASI'] == 588]
            sort589 = df[df['LOKASI'] == 589]
            sort594 = df[df['LOKASI'] == 594]
            sort661 = df[df['LOKASI'] == 661]
            sort662 = df[df['LOKASI'] == 662]
            sort663 = df[df['LOKASI'] == 663]
            sort664 = df[df['LOKASI'] == 664]

            # best150
            best150_all = best150.sort_values(
                by=['RANK NAS.'], ascending=[True])
            del best150_all['LOKASI']
            del best150_all['RANK LOK.']
            best150_all = best150_all.drop(
                best150_all[(best150_all['RANK NAS.'] > 150)].index)

            # 10 besar setiap lokasi
            # 530
            sort530_10 = sort530.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort530_10['LOKASI']
            sort530_10 = sort530_10.drop(
                sort530_10[(sort530_10['RANK LOK.'] > 10)].index)
            # 531
            sort531_10 = sort531.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort531_10['LOKASI']
            sort531_10 = sort531_10.drop(
                sort531_10[(sort531_10['RANK LOK.'] > 10)].index)
            # 532
            sort532_10 = sort532.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort532_10['LOKASI']
            sort532_10 = sort532_10.drop(
                sort532_10[(sort532_10['RANK LOK.'] > 10)].index)
            # 533
            sort533_10 = sort533.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort533_10['LOKASI']
            sort533_10 = sort533_10.drop(
                sort533_10[(sort533_10['RANK LOK.'] > 10)].index)
            # 534
            sort534_10 = sort534.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort534_10['LOKASI']
            sort534_10 = sort534_10.drop(
                sort534_10[(sort534_10['RANK LOK.'] > 10)].index)
            # 535
            sort535_10 = sort535.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort535_10['LOKASI']
            sort535_10 = sort535_10.drop(
                sort535_10[(sort535_10['RANK LOK.'] > 10)].index)
            # 546
            sort546_10 = sort546.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort546_10['LOKASI']
            sort546_10 = sort546_10.drop(
                sort546_10[(sort546_10['RANK LOK.'] > 10)].index)
            # 547
            sort547_10 = sort547.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort547_10['LOKASI']
            sort547_10 = sort547_10.drop(
                sort547_10[(sort547_10['RANK LOK.'] > 10)].index)
            # 548
            sort548_10 = sort548.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort548_10['LOKASI']
            sort548_10 = sort548_10.drop(
                sort548_10[(sort548_10['RANK LOK.'] > 10)].index)
            # 549
            sort549_10 = sort549.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort549_10['LOKASI']
            sort549_10 = sort549_10.drop(
                sort549_10[(sort549_10['RANK LOK.'] > 10)].index)
            # 556
            sort556_10 = sort556.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort556_10['LOKASI']
            sort556_10 = sort556_10.drop(
                sort556_10[(sort556_10['RANK LOK.'] > 10)].index)
            # 557
            sort557_10 = sort557.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort557_10['LOKASI']
            sort557_10 = sort557_10.drop(
                sort557_10[(sort557_10['RANK LOK.'] > 10)].index)
            # 558
            sort558_10 = sort558.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort558_10['LOKASI']
            sort558_10 = sort558_10.drop(
                sort558_10[(sort558_10['RANK LOK.'] > 10)].index)
            # 575
            sort575_10 = sort575.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort575_10['LOKASI']
            sort575_10 = sort575_10.drop(
                sort575_10[(sort575_10['RANK LOK.'] > 10)].index)
            # 576
            sort576_10 = sort576.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort576_10['LOKASI']
            sort576_10 = sort576_10.drop(
                sort576_10[(sort576_10['RANK LOK.'] > 10)].index)
            # 577
            sort577_10 = sort577.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort577_10['LOKASI']
            sort577_10 = sort577_10.drop(
                sort577_10[(sort577_10['RANK LOK.'] > 10)].index)
            # 578
            sort578_10 = sort578.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort578_10['LOKASI']
            sort578_10 = sort578_10.drop(
                sort578_10[(sort578_10['RANK LOK.'] > 10)].index)
            # 588
            sort588_10 = sort588.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort588_10['LOKASI']
            sort588_10 = sort588_10.drop(
                sort588_10[(sort588_10['RANK LOK.'] > 10)].index)
            # 589
            sort589_10 = sort589.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort589_10['LOKASI']
            sort589_10 = sort589_10.drop(
                sort589_10[(sort589_10['RANK LOK.'] > 10)].index)
            # 594
            sort594_10 = sort594.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort594_10['LOKASI']
            sort594_10 = sort594_10.drop(
                sort594_10[(sort594_10['RANK LOK.'] > 10)].index)
            # 661
            sort661_10 = sort661.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort661_10['LOKASI']
            sort661_10 = sort661_10.drop(
                sort661_10[(sort661_10['RANK LOK.'] > 10)].index)
            # 662
            sort662_10 = sort662.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort662_10['LOKASI']
            sort662_10 = sort662_10.drop(
                sort662_10[(sort662_10['RANK LOK.'] > 10)].index)
            # 663
            sort663_10 = sort663.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort663_10['LOKASI']
            sort663_10 = sort663_10.drop(
                sort663_10[(sort663_10['RANK LOK.'] > 10)].index)
            # 664
            sort664_10 = sort664.sort_values(
                by=['RANK LOK.'], ascending=[True])
            del sort664_10['LOKASI']
            sort664_10 = sort664_10.drop(
                sort664_10[(sort664_10['RANK LOK.'] > 10)].index)

            # All 530
            sort530 = sort530.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort530['LOKASI']
            # All 531
            sort531 = sort531.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort531['LOKASI']
            # All 532
            sort532 = sort532.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort532['LOKASI']
            # All 533
            sort533 = sort533.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort533['LOKASI']
            # All 534
            sort534 = sort534.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort534['LOKASI']
            # All 535
            sort535 = sort535.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort535['LOKASI']
            # All 546
            sort546 = sort546.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort546['LOKASI']
            # All 547
            sort547 = sort547.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort547['LOKASI']
            # All 548
            sort548 = sort548.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort548['LOKASI']
            # All 549
            sort549 = sort549.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort549['LOKASI']
            # All 556
            sort556 = sort556.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort556['LOKASI']
            # All 557
            sort557 = sort557.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort557['LOKASI']
            # All 558
            sort558 = sort558.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort558['LOKASI']
            # All 575
            sort575 = sort575.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort575['LOKASI']
            # All 576
            sort576 = sort576.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort576['LOKASI']
            # All 577
            sort577 = sort577.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort577['LOKASI']
            # All 578
            sort578 = sort578.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort578['LOKASI']
            # All 588
            sort588 = sort588.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort588['LOKASI']
            # All 589
            sort589 = sort589.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort589['LOKASI']
            # All 594
            sort594 = sort594.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort594['LOKASI']
            # All 661
            sort661 = sort661.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort661['LOKASI']
            # All 662
            sort662 = sort662.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort662['LOKASI']
            # All 663
            sort663 = sort663.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort663['LOKASI']
            # All 664
            sort664 = sort664.sort_values(by=['NAMA SISWA'], ascending=[True])
            del sort664['LOKASI']

            # jumlah row
            # 150
            rowBest150_all = best150_all.shape[0]
            rowBest150 = best150.shape[0]
            # 530
            row530_10 = sort530_10.shape[0]
            row530 = sort530.shape[0]
            # 531
            row531_10 = sort531_10.shape[0]
            row531 = sort531.shape[0]
            # 532
            row532_10 = sort532_10.shape[0]
            row532 = sort532.shape[0]
            # 533
            row533_10 = sort533_10.shape[0]
            row533 = sort533.shape[0]
            # 534
            row534_10 = sort534_10.shape[0]
            row534 = sort534.shape[0]
            # 535
            row535_10 = sort535_10.shape[0]
            row535 = sort535.shape[0]
            # 546
            row546_10 = sort546_10.shape[0]
            row546 = sort546.shape[0]
            # 547
            row547_10 = sort547_10.shape[0]
            row547 = sort547.shape[0]
            # 548
            row548_10 = sort548_10.shape[0]
            row548 = sort548.shape[0]
            # 549
            row549_10 = sort549_10.shape[0]
            row549 = sort549.shape[0]
            # 556
            row556_10 = sort556_10.shape[0]
            row556 = sort556.shape[0]
            # 557
            row557_10 = sort557_10.shape[0]
            row557 = sort557.shape[0]
            # 558
            row558_10 = sort558_10.shape[0]
            row558 = sort558.shape[0]
            # 575
            row575_10 = sort575_10.shape[0]
            row575 = sort575.shape[0]
            # 576
            row576_10 = sort576_10.shape[0]
            row576 = sort576.shape[0]
            # 577
            row577_10 = sort577_10.shape[0]
            row577 = sort577.shape[0]
            # 578
            row578_10 = sort578_10.shape[0]
            row578 = sort578.shape[0]
            # 588
            row588_10 = sort588_10.shape[0]
            row588 = sort588.shape[0]
            # 589
            row589_10 = sort589_10.shape[0]
            row589 = sort589.shape[0]
            # 594
            row594_10 = sort594_10.shape[0]
            row594 = sort594.shape[0]
            # 661
            row661_10 = sort661_10.shape[0]
            row661 = sort661.shape[0]
            # 662
            row662_10 = sort662_10.shape[0]
            row662 = sort662.shape[0]
            # 663
            row663_10 = sort663_10.shape[0]
            row663 = sort663.shape[0]
            # 664
            row664_10 = sort664_10.shape[0]
            row664 = sort664.shape[0]

            # Create a Pandas Excel writer using XlsxWriter as the engine.
            # Path file hasil penyimpanan
            file_name = f"{kelas}_{penilaian}_{semester}_lokasi_depok_padang.xlsx"
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

            # 530
            # Convert the dataframe to an XlsxWriter Excel object.
            sort530_10.to_excel(writer, sheet_name='530',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort530.to_excel(writer, sheet_name='530',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 531
            # Convert the dataframe to an XlsxWriter Excel object.
            sort531_10.to_excel(writer, sheet_name='531',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort531.to_excel(writer, sheet_name='531',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 532
            # Convert the dataframe to an XlsxWriter Excel object.
            sort532_10.to_excel(writer, sheet_name='532',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort532.to_excel(writer, sheet_name='532',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 533
            # Convert the dataframe to an XlsxWriter Excel object.
            sort533_10.to_excel(writer, sheet_name='533',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort533.to_excel(writer, sheet_name='533',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 534
            # Convert the dataframe to an XlsxWriter Excel object.
            sort534_10.to_excel(writer, sheet_name='534',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort534.to_excel(writer, sheet_name='534',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 535
            # Convert the dataframe to an XlsxWriter Excel object.
            sort535_10.to_excel(writer, sheet_name='535',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort535.to_excel(writer, sheet_name='535',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 546
            # Convert the dataframe to an XlsxWriter Excel object.
            sort546_10.to_excel(writer, sheet_name='546',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort546.to_excel(writer, sheet_name='546',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 547
            # Convert the dataframe to an XlsxWriter Excel object.
            sort547_10.to_excel(writer, sheet_name='547',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort547.to_excel(writer, sheet_name='547',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 548
            # Convert the dataframe to an XlsxWriter Excel object.
            sort548_10.to_excel(writer, sheet_name='548',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort548.to_excel(writer, sheet_name='548',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 549
            # Convert the dataframe to an XlsxWriter Excel object.
            sort549_10.to_excel(writer, sheet_name='549',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort549.to_excel(writer, sheet_name='549',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 556
            # Convert the dataframe to an XlsxWriter Excel object.
            sort556_10.to_excel(writer, sheet_name='556',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort556.to_excel(writer, sheet_name='556',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 557
            # Convert the dataframe to an XlsxWriter Excel object.
            sort557_10.to_excel(writer, sheet_name='557',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort557.to_excel(writer, sheet_name='557',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 558
            # Convert the dataframe to an XlsxWriter Excel object.
            sort558_10.to_excel(writer, sheet_name='558',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort558.to_excel(writer, sheet_name='558',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 575
            # Convert the dataframe to an XlsxWriter Excel object.
            sort575_10.to_excel(writer, sheet_name='575',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort575.to_excel(writer, sheet_name='575',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 576
            # Convert the dataframe to an XlsxWriter Excel object.
            sort576_10.to_excel(writer, sheet_name='576',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort576.to_excel(writer, sheet_name='576',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 577
            # Convert the dataframe to an XlsxWriter Excel object.
            sort577_10.to_excel(writer, sheet_name='577',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort577.to_excel(writer, sheet_name='577',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 578
            # Convert the dataframe to an XlsxWriter Excel object.
            sort578_10.to_excel(writer, sheet_name='578',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort578.to_excel(writer, sheet_name='578',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 588
            # Convert the dataframe to an XlsxWriter Excel object.
            sort588_10.to_excel(writer, sheet_name='588',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort588.to_excel(writer, sheet_name='588',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 589
            # Convert the dataframe to an XlsxWriter Excel object.
            sort589_10.to_excel(writer, sheet_name='589',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort589.to_excel(writer, sheet_name='589',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 594
            # Convert the dataframe to an XlsxWriter Excel object.
            sort594_10.to_excel(writer, sheet_name='594',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort594.to_excel(writer, sheet_name='594',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 661
            # Convert the dataframe to an XlsxWriter Excel object.
            sort661_10.to_excel(writer, sheet_name='661',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort661.to_excel(writer, sheet_name='661',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 662
            # Convert the dataframe to an XlsxWriter Excel object.
            sort662_10.to_excel(writer, sheet_name='662',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort662.to_excel(writer, sheet_name='662',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 663
            # Convert the dataframe to an XlsxWriter Excel object.
            sort663_10.to_excel(writer, sheet_name='663',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort663.to_excel(writer, sheet_name='663',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)
            # 664
            # Convert the dataframe to an XlsxWriter Excel object.
            sort664_10.to_excel(writer, sheet_name='664',
                                startrow=5,
                                startcol=0,
                                index=False,
                                header=False)
            # Convert the dataframe to an XlsxWriter Excel object.
            sort664.to_excel(writer, sheet_name='664',
                             startrow=22,
                             startcol=0,
                             index=False,
                             header=False)

            # Get the xlsxwriter objects from the dataframe writer object.
            workbook = writer.book

            # membuat worksheet baru
            worksheetcover = writer.sheets['cover']
            worksheetbest = writer.sheets['best_150']
            worksheet530 = writer.sheets['530']
            worksheet531 = writer.sheets['531']
            worksheet532 = writer.sheets['532']
            worksheet533 = writer.sheets['533']
            worksheet534 = writer.sheets['534']
            worksheet535 = writer.sheets['535']
            worksheet546 = writer.sheets['546']
            worksheet547 = writer.sheets['547']
            worksheet548 = writer.sheets['548']
            worksheet549 = writer.sheets['549']
            worksheet556 = writer.sheets['556']
            worksheet557 = writer.sheets['557']
            worksheet558 = writer.sheets['558']
            worksheet575 = writer.sheets['575']
            worksheet576 = writer.sheets['576']
            worksheet577 = writer.sheets['577']
            worksheet578 = writer.sheets['578']
            worksheet588 = writer.sheets['588']
            worksheet589 = writer.sheets['589']
            worksheet594 = writer.sheets['594']
            worksheet661 = writer.sheets['661']
            worksheet662 = writer.sheets['662']
            worksheet663 = writer.sheets['663']
            worksheet664 = writer.sheets['664']

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
                'A1:Q1', fr'150 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF NASIONAL', title)
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

            # worksheet 530
            worksheet530.insert_image('A1', r'logo resmi nf.jpg')

            worksheet530.set_column('A:A', 7, center)
            worksheet530.set_column('B:B', 6, center)
            worksheet530.set_column('C:C', 18.14, center)
            worksheet530.set_column('D:D', 25, left)
            worksheet530.set_column('E:E', 13.14, left)
            worksheet530.set_column('F:F', 8.57, center)
            worksheet530.set_column('G:R', 5, center)
            worksheet530.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF POLSEK DEPOK', title)
            worksheet530.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet530.write('A5', 'LOKASI', header)
            worksheet530.write('B5', 'TOTAL', header)
            worksheet530.merge_range('A4:B4', 'RANK', header)
            worksheet530.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet530.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet530.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet530.merge_range('F4:F5', 'KELAS', header)
            worksheet530.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet530.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet530.write('G5', 'MAT', body)
            worksheet530.write('H5', 'IND', body)
            worksheet530.write('I5', 'ENG', body)
            worksheet530.write('J5', 'IPA', body)
            worksheet530.write('K5', 'IPS', body)
            worksheet530.write('L5', 'JML', body)
            worksheet530.write('M5', 'MAT', body)
            worksheet530.write('N5', 'IND', body)
            worksheet530.write('O5', 'ENG', body)
            worksheet530.write('P5', 'IPA', body)
            worksheet530.write('Q5', 'IPS', body)
            worksheet530.write('R5', 'JML', body)

            worksheet530.conditional_format(5, 0, row530_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet530.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF POLSEK DEPOK', title)
            worksheet530.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet530.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet530.write('A22', 'LOKASI', header)
            worksheet530.write('B22', 'TOTAL', header)
            worksheet530.merge_range('A21:B21', 'RANK', header)
            worksheet530.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet530.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet530.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet530.merge_range('F21:F22', 'KELAS', header)
            worksheet530.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet530.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet530.write('G22', 'MAT', body)
            worksheet530.write('H22', 'IND', body)
            worksheet530.write('I22', 'ENG', body)
            worksheet530.write('J22', 'IPA', body)
            worksheet530.write('K22', 'IPS', body)
            worksheet530.write('L22', 'JML', body)
            worksheet530.write('M22', 'MAT', body)
            worksheet530.write('N22', 'IND', body)
            worksheet530.write('O22', 'ENG', body)
            worksheet530.write('P22', 'IPA', body)
            worksheet530.write('Q22', 'IPS', body)
            worksheet530.write('R22', 'JML', body)

            worksheet530.conditional_format(22, 0, row530+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 531
            worksheet531.insert_image('A1', r'logo resmi nf.jpg')

            worksheet531.set_column('A:A', 7, center)
            worksheet531.set_column('B:B', 6, center)
            worksheet531.set_column('C:C', 18.14, center)
            worksheet531.set_column('D:D', 25, left)
            worksheet531.set_column('E:E', 13.14, left)
            worksheet531.set_column('F:F', 8.57, center)
            worksheet531.set_column('G:R', 5, center)
            worksheet531.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF DEPOK 1', title)
            worksheet531.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet531.write('A5', 'LOKASI', header)
            worksheet531.write('B5', 'TOTAL', header)
            worksheet531.merge_range('A4:B4', 'RANK', header)
            worksheet531.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet531.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet531.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet531.merge_range('F4:F5', 'KELAS', header)
            worksheet531.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet531.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet531.write('G5', 'MAT', body)
            worksheet531.write('H5', 'IND', body)
            worksheet531.write('I5', 'ENG', body)
            worksheet531.write('J5', 'IPA', body)
            worksheet531.write('K5', 'IPS', body)
            worksheet531.write('L5', 'JML', body)
            worksheet531.write('M5', 'MAT', body)
            worksheet531.write('N5', 'IND', body)
            worksheet531.write('O5', 'ENG', body)
            worksheet531.write('P5', 'IPA', body)
            worksheet531.write('Q5', 'IPS', body)
            worksheet531.write('R5', 'JML', body)

            worksheet531.conditional_format(5, 0, row531_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet531.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF DEPOK 1', title)
            worksheet531.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet531.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet531.write('A22', 'LOKASI', header)
            worksheet531.write('B22', 'TOTAL', header)
            worksheet531.merge_range('A21:B21', 'RANK', header)
            worksheet531.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet531.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet531.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet531.merge_range('F21:F22', 'KELAS', header)
            worksheet531.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet531.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet531.write('G22', 'MAT', body)
            worksheet531.write('H22', 'IND', body)
            worksheet531.write('I22', 'ENG', body)
            worksheet531.write('J22', 'IPA', body)
            worksheet531.write('K22', 'IPS', body)
            worksheet531.write('L22', 'JML', body)
            worksheet531.write('M22', 'MAT', body)
            worksheet531.write('N22', 'IND', body)
            worksheet531.write('O22', 'ENG', body)
            worksheet531.write('P22', 'IPA', body)
            worksheet531.write('Q22', 'IPS', body)
            worksheet531.write('R22', 'JML', body)

            worksheet531.conditional_format(22, 0, row531+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 532
            worksheet532.insert_image('A1', r'logo resmi nf.jpg')

            worksheet532.set_column('A:A', 7, center)
            worksheet532.set_column('B:B', 6, center)
            worksheet532.set_column('C:C', 18.14, center)
            worksheet532.set_column('D:D', 25, left)
            worksheet532.set_column('E:E', 13.14, left)
            worksheet532.set_column('F:F', 8.57, center)
            worksheet532.set_column('G:R', 5, center)
            worksheet532.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PROKLAMASI DEPOK 2', title)
            worksheet532.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet532.write('A5', 'LOKASI', header)
            worksheet532.write('B5', 'TOTAL', header)
            worksheet532.merge_range('A4:B4', 'RANK', header)
            worksheet532.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet532.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet532.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet532.merge_range('F4:F5', 'KELAS', header)
            worksheet532.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet532.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet532.write('G5', 'MAT', body)
            worksheet532.write('H5', 'IND', body)
            worksheet532.write('I5', 'ENG', body)
            worksheet532.write('J5', 'IPA', body)
            worksheet532.write('K5', 'IPS', body)
            worksheet532.write('L5', 'JML', body)
            worksheet532.write('M5', 'MAT', body)
            worksheet532.write('N5', 'IND', body)
            worksheet532.write('O5', 'ENG', body)
            worksheet532.write('P5', 'IPA', body)
            worksheet532.write('Q5', 'IPS', body)
            worksheet532.write('R5', 'JML', body)

            worksheet532.conditional_format(5, 0, row532_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet532.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF PROKLAMASI DEPOK 2', title)
            worksheet532.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet532.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet532.write('A22', 'LOKASI', header)
            worksheet532.write('B22', 'TOTAL', header)
            worksheet532.merge_range('A21:B21', 'RANK', header)
            worksheet532.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet532.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet532.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet532.merge_range('F21:F22', 'KELAS', header)
            worksheet532.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet532.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet532.write('G22', 'MAT', body)
            worksheet532.write('H22', 'IND', body)
            worksheet532.write('I22', 'ENG', body)
            worksheet532.write('J22', 'IPA', body)
            worksheet532.write('K22', 'IPS', body)
            worksheet532.write('L22', 'JML', body)
            worksheet532.write('M22', 'MAT', body)
            worksheet532.write('N22', 'IND', body)
            worksheet532.write('O22', 'ENG', body)
            worksheet532.write('P22', 'IPA', body)
            worksheet532.write('Q22', 'IPS', body)
            worksheet532.write('R22', 'JML', body)

            worksheet532.conditional_format(22, 0, row532+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 533
            worksheet533.insert_image('A1', r'logo resmi nf.jpg')

            worksheet533.set_column('A:A', 7, center)
            worksheet533.set_column('B:B', 6, center)
            worksheet533.set_column('C:C', 18.14, center)
            worksheet533.set_column('D:D', 25, left)
            worksheet533.set_column('E:E', 13.14, left)
            worksheet533.set_column('F:F', 8.57, center)
            worksheet533.set_column('G:R', 5, center)
            worksheet533.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIMANGGIS', title)
            worksheet533.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet533.write('A5', 'LOKASI', header)
            worksheet533.write('B5', 'TOTAL', header)
            worksheet533.merge_range('A4:B4', 'RANK', header)
            worksheet533.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet533.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet533.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet533.merge_range('F4:F5', 'KELAS', header)
            worksheet533.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet533.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet533.write('G5', 'MAT', body)
            worksheet533.write('H5', 'IND', body)
            worksheet533.write('I5', 'ENG', body)
            worksheet533.write('J5', 'IPA', body)
            worksheet533.write('K5', 'IPS', body)
            worksheet533.write('L5', 'JML', body)
            worksheet533.write('M5', 'MAT', body)
            worksheet533.write('N5', 'IND', body)
            worksheet533.write('O5', 'ENG', body)
            worksheet533.write('P5', 'IPA', body)
            worksheet533.write('Q5', 'IPS', body)
            worksheet533.write('R5', 'JML', body)

            worksheet533.conditional_format(5, 0, row533_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet533.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF CIMANGGIS', title)
            worksheet533.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet533.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet533.write('A22', 'LOKASI', header)
            worksheet533.write('B22', 'TOTAL', header)
            worksheet533.merge_range('A21:B21', 'RANK', header)
            worksheet533.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet533.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet533.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet533.merge_range('F21:F22', 'KELAS', header)
            worksheet533.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet533.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet533.write('G22', 'MAT', body)
            worksheet533.write('H22', 'IND', body)
            worksheet533.write('I22', 'ENG', body)
            worksheet533.write('J22', 'IPA', body)
            worksheet533.write('K22', 'IPS', body)
            worksheet533.write('L22', 'JML', body)
            worksheet533.write('M22', 'MAT', body)
            worksheet533.write('N22', 'IND', body)
            worksheet533.write('O22', 'ENG', body)
            worksheet533.write('P22', 'IPA', body)
            worksheet533.write('Q22', 'IPS', body)
            worksheet533.write('R22', 'JML', body)

            worksheet533.conditional_format(22, 0, row533+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 534
            worksheet534.insert_image('A1', r'logo resmi nf.jpg')

            worksheet534.set_column('A:A', 7, center)
            worksheet534.set_column('B:B', 6, center)
            worksheet534.set_column('C:C', 18.14, center)
            worksheet534.set_column('D:D', 25, left)
            worksheet534.set_column('E:E', 13.14, left)
            worksheet534.set_column('F:F', 8.57, center)
            worksheet534.set_column('G:R', 5, center)
            worksheet534.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SAWANGAN', title)
            worksheet534.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet534.write('A5', 'LOKASI', header)
            worksheet534.write('B5', 'TOTAL', header)
            worksheet534.merge_range('A4:B4', 'RANK', header)
            worksheet534.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet534.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet534.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet534.merge_range('F4:F5', 'KELAS', header)
            worksheet534.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet534.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet534.write('G5', 'MAT', body)
            worksheet534.write('H5', 'IND', body)
            worksheet534.write('I5', 'ENG', body)
            worksheet534.write('J5', 'IPA', body)
            worksheet534.write('K5', 'IPS', body)
            worksheet534.write('L5', 'JML', body)
            worksheet534.write('M5', 'MAT', body)
            worksheet534.write('N5', 'IND', body)
            worksheet534.write('O5', 'ENG', body)
            worksheet534.write('P5', 'IPA', body)
            worksheet534.write('Q5', 'IPS', body)
            worksheet534.write('R5', 'JML', body)

            worksheet534.conditional_format(5, 0, row534_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet534.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF SAWANGAN', title)
            worksheet534.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet534.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet534.write('A22', 'LOKASI', header)
            worksheet534.write('B22', 'TOTAL', header)
            worksheet534.merge_range('A21:B21', 'RANK', header)
            worksheet534.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet534.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet534.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet534.merge_range('F21:F22', 'KELAS', header)
            worksheet534.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet534.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet534.write('G22', 'MAT', body)
            worksheet534.write('H22', 'IND', body)
            worksheet534.write('I22', 'ENG', body)
            worksheet534.write('J22', 'IPA', body)
            worksheet534.write('K22', 'IPS', body)
            worksheet534.write('L22', 'JML', body)
            worksheet534.write('M22', 'MAT', body)
            worksheet534.write('N22', 'IND', body)
            worksheet534.write('O22', 'ENG', body)
            worksheet534.write('P22', 'IPA', body)
            worksheet534.write('Q22', 'IPS', body)
            worksheet534.write('R22', 'JML', body)

            worksheet534.conditional_format(22, 0, row534+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 535
            worksheet535.insert_image('A1', r'logo resmi nf.jpg')

            worksheet535.set_column('A:A', 7, center)
            worksheet535.set_column('B:B', 6, center)
            worksheet535.set_column('C:C', 18.14, center)
            worksheet535.set_column('D:D', 25, left)
            worksheet535.set_column('E:E', 13.14, left)
            worksheet535.set_column('F:F', 8.57, center)
            worksheet535.set_column('G:R', 5, center)
            worksheet535.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF BRIMOB', title)
            worksheet535.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet535.write('A5', 'LOKASI', header)
            worksheet535.write('B5', 'TOTAL', header)
            worksheet535.merge_range('A4:B4', 'RANK', header)
            worksheet535.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet535.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet535.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet535.merge_range('F4:F5', 'KELAS', header)
            worksheet535.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet535.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet535.write('G5', 'MAT', body)
            worksheet535.write('H5', 'IND', body)
            worksheet535.write('I5', 'ENG', body)
            worksheet535.write('J5', 'IPA', body)
            worksheet535.write('K5', 'IPS', body)
            worksheet535.write('L5', 'JML', body)
            worksheet535.write('M5', 'MAT', body)
            worksheet535.write('N5', 'IND', body)
            worksheet535.write('O5', 'ENG', body)
            worksheet535.write('P5', 'IPA', body)
            worksheet535.write('Q5', 'IPS', body)
            worksheet535.write('R5', 'JML', body)

            worksheet535.conditional_format(5, 0, row535_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet535.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF BRIMOB', title)
            worksheet535.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet535.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet535.write('A22', 'LOKASI', header)
            worksheet535.write('B22', 'TOTAL', header)
            worksheet535.merge_range('A21:B21', 'RANK', header)
            worksheet535.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet535.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet535.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet535.merge_range('F21:F22', 'KELAS', header)
            worksheet535.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet535.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet535.write('G22', 'MAT', body)
            worksheet535.write('H22', 'IND', body)
            worksheet535.write('I22', 'ENG', body)
            worksheet535.write('J22', 'IPA', body)
            worksheet535.write('K22', 'IPS', body)
            worksheet535.write('L22', 'JML', body)
            worksheet535.write('M22', 'MAT', body)
            worksheet535.write('N22', 'IND', body)
            worksheet535.write('O22', 'ENG', body)
            worksheet535.write('P22', 'IPA', body)
            worksheet535.write('Q22', 'IPS', body)
            worksheet535.write('R22', 'JML', body)

            worksheet535.conditional_format(22, 0, row535+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 546
            worksheet546.insert_image('A1', r'logo resmi nf.jpg')

            worksheet546.set_column('A:A', 7, center)
            worksheet546.set_column('B:B', 6, center)
            worksheet546.set_column('C:C', 18.14, center)
            worksheet546.set_column('D:D', 25, left)
            worksheet546.set_column('E:E', 13.14, left)
            worksheet546.set_column('F:F', 8.57, center)
            worksheet546.set_column('G:R', 5, center)
            worksheet546.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PAJAJARAN (PPIB)', title)
            worksheet546.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet546.write('A5', 'LOKASI', header)
            worksheet546.write('B5', 'TOTAL', header)
            worksheet546.merge_range('A4:B4', 'RANK', header)
            worksheet546.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet546.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet546.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet546.merge_range('F4:F5', 'KELAS', header)
            worksheet546.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet546.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet546.write('G5', 'MAT', body)
            worksheet546.write('H5', 'IND', body)
            worksheet546.write('I5', 'ENG', body)
            worksheet546.write('J5', 'IPA', body)
            worksheet546.write('K5', 'IPS', body)
            worksheet546.write('L5', 'JML', body)
            worksheet546.write('M5', 'MAT', body)
            worksheet546.write('N5', 'IND', body)
            worksheet546.write('O5', 'ENG', body)
            worksheet546.write('P5', 'IPA', body)
            worksheet546.write('Q5', 'IPS', body)
            worksheet546.write('R5', 'JML', body)

            worksheet546.conditional_format(5, 0, row546_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet546.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF PAJAJARAN (PPIB)', title)
            worksheet546.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet546.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet546.write('A22', 'LOKASI', header)
            worksheet546.write('B22', 'TOTAL', header)
            worksheet546.merge_range('A21:B21', 'RANK', header)
            worksheet546.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet546.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet546.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet546.merge_range('F21:F22', 'KELAS', header)
            worksheet546.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet546.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet546.write('G22', 'MAT', body)
            worksheet546.write('H22', 'IND', body)
            worksheet546.write('I22', 'ENG', body)
            worksheet546.write('J22', 'IPA', body)
            worksheet546.write('K22', 'IPS', body)
            worksheet546.write('L22', 'JML', body)
            worksheet546.write('M22', 'MAT', body)
            worksheet546.write('N22', 'IND', body)
            worksheet546.write('O22', 'ENG', body)
            worksheet546.write('P22', 'IPA', body)
            worksheet546.write('Q22', 'IPS', body)
            worksheet546.write('R22', 'JML', body)

            worksheet546.conditional_format(22, 0, row546+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 547
            worksheet547.insert_image('A1', r'logo resmi nf.jpg')

            worksheet547.set_column('A:A', 7, center)
            worksheet547.set_column('B:B', 6, center)
            worksheet547.set_column('C:C', 18.14, center)
            worksheet547.set_column('D:D', 25, left)
            worksheet547.set_column('E:E', 13.14, left)
            worksheet547.set_column('F:F', 8.57, center)
            worksheet547.set_column('G:R', 5, center)
            worksheet547.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SUKASARI', title)
            worksheet547.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet547.write('A5', 'LOKASI', header)
            worksheet547.write('B5', 'TOTAL', header)
            worksheet547.merge_range('A4:B4', 'RANK', header)
            worksheet547.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet547.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet547.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet547.merge_range('F4:F5', 'KELAS', header)
            worksheet547.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet547.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet547.write('G5', 'MAT', body)
            worksheet547.write('H5', 'IND', body)
            worksheet547.write('I5', 'ENG', body)
            worksheet547.write('J5', 'IPA', body)
            worksheet547.write('K5', 'IPS', body)
            worksheet547.write('L5', 'JML', body)
            worksheet547.write('M5', 'MAT', body)
            worksheet547.write('N5', 'IND', body)
            worksheet547.write('O5', 'ENG', body)
            worksheet547.write('P5', 'IPA', body)
            worksheet547.write('Q5', 'IPS', body)
            worksheet547.write('R5', 'JML', body)

            worksheet547.conditional_format(5, 0, row547_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet547.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF SUKASARI', title)
            worksheet547.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet547.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet547.write('A22', 'LOKASI', header)
            worksheet547.write('B22', 'TOTAL', header)
            worksheet547.merge_range('A21:B21', 'RANK', header)
            worksheet547.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet547.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet547.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet547.merge_range('F21:F22', 'KELAS', header)
            worksheet547.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet547.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet547.write('G22', 'MAT', body)
            worksheet547.write('H22', 'IND', body)
            worksheet547.write('I22', 'ENG', body)
            worksheet547.write('J22', 'IPA', body)
            worksheet547.write('K22', 'IPS', body)
            worksheet547.write('L22', 'JML', body)
            worksheet547.write('M22', 'MAT', body)
            worksheet547.write('N22', 'IND', body)
            worksheet547.write('O22', 'ENG', body)
            worksheet547.write('P22', 'IPA', body)
            worksheet547.write('Q22', 'IPS', body)
            worksheet547.write('R22', 'JML', body)

            worksheet547.conditional_format(22, 0, row547+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 548
            worksheet548.insert_image('A1', r'logo resmi nf.jpg')

            worksheet548.set_column('A:A', 7, center)
            worksheet548.set_column('B:B', 6, center)
            worksheet548.set_column('C:C', 18.14, center)
            worksheet548.set_column('D:D', 25, left)
            worksheet548.set_column('E:E', 13.14, left)
            worksheet548.set_column('F:F', 8.57, center)
            worksheet548.set_column('G:R', 5, center)
            worksheet548.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF BUBULAK', title)
            worksheet548.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet548.write('A5', 'LOKASI', header)
            worksheet548.write('B5', 'TOTAL', header)
            worksheet548.merge_range('A4:B4', 'RANK', header)
            worksheet548.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet548.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet548.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet548.merge_range('F4:F5', 'KELAS', header)
            worksheet548.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet548.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet548.write('G5', 'MAT', body)
            worksheet548.write('H5', 'IND', body)
            worksheet548.write('I5', 'ENG', body)
            worksheet548.write('J5', 'IPA', body)
            worksheet548.write('K5', 'IPS', body)
            worksheet548.write('L5', 'JML', body)
            worksheet548.write('M5', 'MAT', body)
            worksheet548.write('N5', 'IND', body)
            worksheet548.write('O5', 'ENG', body)
            worksheet548.write('P5', 'IPA', body)
            worksheet548.write('Q5', 'IPS', body)
            worksheet548.write('R5', 'JML', body)

            worksheet548.conditional_format(5, 0, row548_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet548.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF BUBULAK', title)
            worksheet548.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet548.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet548.write('A22', 'LOKASI', header)
            worksheet548.write('B22', 'TOTAL', header)
            worksheet548.merge_range('A21:B21', 'RANK', header)
            worksheet548.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet548.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet548.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet548.merge_range('F21:F22', 'KELAS', header)
            worksheet548.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet548.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet548.write('G22', 'MAT', body)
            worksheet548.write('H22', 'IND', body)
            worksheet548.write('I22', 'ENG', body)
            worksheet548.write('J22', 'IPA', body)
            worksheet548.write('K22', 'IPS', body)
            worksheet548.write('L22', 'JML', body)
            worksheet548.write('M22', 'MAT', body)
            worksheet548.write('N22', 'IND', body)
            worksheet548.write('O22', 'ENG', body)
            worksheet548.write('P22', 'IPA', body)
            worksheet548.write('Q22', 'IPS', body)
            worksheet548.write('R22', 'JML', body)

            worksheet548.conditional_format(22, 0, row548+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 549
            worksheet549.insert_image('A1', r'logo resmi nf.jpg')

            worksheet549.set_column('A:A', 7, center)
            worksheet549.set_column('B:B', 6, center)
            worksheet549.set_column('C:C', 18.14, center)
            worksheet549.set_column('D:D', 25, left)
            worksheet549.set_column('E:E', 13.14, left)
            worksheet549.set_column('F:F', 8.57, center)
            worksheet549.set_column('G:R', 5, center)
            worksheet549.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CITEUREUP', title)
            worksheet549.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet549.write('A5', 'LOKASI', header)
            worksheet549.write('B5', 'TOTAL', header)
            worksheet549.merge_range('A4:B4', 'RANK', header)
            worksheet549.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet549.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet549.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet549.merge_range('F4:F5', 'KELAS', header)
            worksheet549.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet549.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet549.write('G5', 'MAT', body)
            worksheet549.write('H5', 'IND', body)
            worksheet549.write('I5', 'ENG', body)
            worksheet549.write('J5', 'IPA', body)
            worksheet549.write('K5', 'IPS', body)
            worksheet549.write('L5', 'JML', body)
            worksheet549.write('M5', 'MAT', body)
            worksheet549.write('N5', 'IND', body)
            worksheet549.write('O5', 'ENG', body)
            worksheet549.write('P5', 'IPA', body)
            worksheet549.write('Q5', 'IPS', body)
            worksheet549.write('R5', 'JML', body)

            worksheet549.conditional_format(5, 0, row549_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet549.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF CITEUREUP', title)
            worksheet549.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet549.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet549.write('A22', 'LOKASI', header)
            worksheet549.write('B22', 'TOTAL', header)
            worksheet549.merge_range('A21:B21', 'RANK', header)
            worksheet549.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet549.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet549.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet549.merge_range('F21:F22', 'KELAS', header)
            worksheet549.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet549.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet549.write('G22', 'MAT', body)
            worksheet549.write('H22', 'IND', body)
            worksheet549.write('I22', 'ENG', body)
            worksheet549.write('J22', 'IPA', body)
            worksheet549.write('K22', 'IPS', body)
            worksheet549.write('L22', 'JML', body)
            worksheet549.write('M22', 'MAT', body)
            worksheet549.write('N22', 'IND', body)
            worksheet549.write('O22', 'ENG', body)
            worksheet549.write('P22', 'IPA', body)
            worksheet549.write('Q22', 'IPS', body)
            worksheet549.write('R22', 'JML', body)

            worksheet549.conditional_format(22, 0, row549+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 556
            worksheet556.insert_image('A1', r'logo resmi nf.jpg')

            worksheet556.set_column('A:A', 7, center)
            worksheet556.set_column('B:B', 6, center)
            worksheet556.set_column('C:C', 18.14, center)
            worksheet556.set_column('D:D', 25, left)
            worksheet556.set_column('E:E', 13.14, left)
            worksheet556.set_column('F:F', 8.57, center)
            worksheet556.set_column('G:R', 5, center)
            worksheet556.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF YASMIN', title)
            worksheet556.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet556.write('A5', 'LOKASI', header)
            worksheet556.write('B5', 'TOTAL', header)
            worksheet556.merge_range('A4:B4', 'RANK', header)
            worksheet556.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet556.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet556.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet556.merge_range('F4:F5', 'KELAS', header)
            worksheet556.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet556.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet556.write('G5', 'MAT', body)
            worksheet556.write('H5', 'IND', body)
            worksheet556.write('I5', 'ENG', body)
            worksheet556.write('J5', 'IPA', body)
            worksheet556.write('K5', 'IPS', body)
            worksheet556.write('L5', 'JML', body)
            worksheet556.write('M5', 'MAT', body)
            worksheet556.write('N5', 'IND', body)
            worksheet556.write('O5', 'ENG', body)
            worksheet556.write('P5', 'IPA', body)
            worksheet556.write('Q5', 'IPS', body)
            worksheet556.write('R5', 'JML', body)

            worksheet556.conditional_format(5, 0, row556_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet556.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF YASMIN', title)
            worksheet556.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet556.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet556.write('A22', 'LOKASI', header)
            worksheet556.write('B22', 'TOTAL', header)
            worksheet556.merge_range('A21:B21', 'RANK', header)
            worksheet556.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet556.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet556.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet556.merge_range('F21:F22', 'KELAS', header)
            worksheet556.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet556.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet556.write('G22', 'MAT', body)
            worksheet556.write('H22', 'IND', body)
            worksheet556.write('I22', 'ENG', body)
            worksheet556.write('J22', 'IPA', body)
            worksheet556.write('K22', 'IPS', body)
            worksheet556.write('L22', 'JML', body)
            worksheet556.write('M22', 'MAT', body)
            worksheet556.write('N22', 'IND', body)
            worksheet556.write('O22', 'ENG', body)
            worksheet556.write('P22', 'IPA', body)
            worksheet556.write('Q22', 'IPS', body)
            worksheet556.write('R22', 'JML', body)

            worksheet556.conditional_format(22, 0, row556+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 557
            worksheet557.insert_image('A1', r'logo resmi nf.jpg')

            worksheet557.set_column('A:A', 7, center)
            worksheet557.set_column('B:B', 6, center)
            worksheet557.set_column('C:C', 18.14, center)
            worksheet557.set_column('D:D', 25, left)
            worksheet557.set_column('E:E', 13.14, left)
            worksheet557.set_column('F:F', 8.57, center)
            worksheet557.set_column('G:R', 5, center)
            worksheet557.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF YOGYA PLAZA', title)
            worksheet557.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet557.write('A5', 'LOKASI', header)
            worksheet557.write('B5', 'TOTAL', header)
            worksheet557.merge_range('A4:B4', 'RANK', header)
            worksheet557.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet557.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet557.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet557.merge_range('F4:F5', 'KELAS', header)
            worksheet557.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet557.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet557.write('G5', 'MAT', body)
            worksheet557.write('H5', 'IND', body)
            worksheet557.write('I5', 'ENG', body)
            worksheet557.write('J5', 'IPA', body)
            worksheet557.write('K5', 'IPS', body)
            worksheet557.write('L5', 'JML', body)
            worksheet557.write('M5', 'MAT', body)
            worksheet557.write('N5', 'IND', body)
            worksheet557.write('O5', 'ENG', body)
            worksheet557.write('P5', 'IPA', body)
            worksheet557.write('Q5', 'IPS', body)
            worksheet557.write('R5', 'JML', body)

            worksheet557.conditional_format(5, 0, row557_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet557.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF YOGYA PLAZA', title)
            worksheet557.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet557.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet557.write('A22', 'LOKASI', header)
            worksheet557.write('B22', 'TOTAL', header)
            worksheet557.merge_range('A21:B21', 'RANK', header)
            worksheet557.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet557.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet557.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet557.merge_range('F21:F22', 'KELAS', header)
            worksheet557.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet557.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet557.write('G22', 'MAT', body)
            worksheet557.write('H22', 'IND', body)
            worksheet557.write('I22', 'ENG', body)
            worksheet557.write('J22', 'IPA', body)
            worksheet557.write('K22', 'IPS', body)
            worksheet557.write('L22', 'JML', body)
            worksheet557.write('M22', 'MAT', body)
            worksheet557.write('N22', 'IND', body)
            worksheet557.write('O22', 'ENG', body)
            worksheet557.write('P22', 'IPA', body)
            worksheet557.write('Q22', 'IPS', body)
            worksheet557.write('R22', 'JML', body)

            worksheet557.conditional_format(22, 0, row557+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 558
            worksheet558.insert_image('A1', r'logo resmi nf.jpg')

            worksheet558.set_column('A:A', 7, center)
            worksheet558.set_column('B:B', 6, center)
            worksheet558.set_column('C:C', 18.14, center)
            worksheet558.set_column('D:D', 25, left)
            worksheet558.set_column('E:E', 13.14, left)
            worksheet558.set_column('F:F', 8.57, center)
            worksheet558.set_column('G:R', 5, center)
            worksheet558.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PARUNG', title)
            worksheet558.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet558.write('A5', 'LOKASI', header)
            worksheet558.write('B5', 'TOTAL', header)
            worksheet558.merge_range('A4:B4', 'RANK', header)
            worksheet558.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet558.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet558.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet558.merge_range('F4:F5', 'KELAS', header)
            worksheet558.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet558.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet558.write('G5', 'MAT', body)
            worksheet558.write('H5', 'IND', body)
            worksheet558.write('I5', 'ENG', body)
            worksheet558.write('J5', 'IPA', body)
            worksheet558.write('K5', 'IPS', body)
            worksheet558.write('L5', 'JML', body)
            worksheet558.write('M5', 'MAT', body)
            worksheet558.write('N5', 'IND', body)
            worksheet558.write('O5', 'ENG', body)
            worksheet558.write('P5', 'IPA', body)
            worksheet558.write('Q5', 'IPS', body)
            worksheet558.write('R5', 'JML', body)

            worksheet558.conditional_format(5, 0, row558_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet558.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF PARUNG', title)
            worksheet558.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet558.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet558.write('A22', 'LOKASI', header)
            worksheet558.write('B22', 'TOTAL', header)
            worksheet558.merge_range('A21:B21', 'RANK', header)
            worksheet558.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet558.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet558.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet558.merge_range('F21:F22', 'KELAS', header)
            worksheet558.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet558.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet558.write('G22', 'MAT', body)
            worksheet558.write('H22', 'IND', body)
            worksheet558.write('I22', 'ENG', body)
            worksheet558.write('J22', 'IPA', body)
            worksheet558.write('K22', 'IPS', body)
            worksheet558.write('L22', 'JML', body)
            worksheet558.write('M22', 'MAT', body)
            worksheet558.write('N22', 'IND', body)
            worksheet558.write('O22', 'ENG', body)
            worksheet558.write('P22', 'IPA', body)
            worksheet558.write('Q22', 'IPS', body)
            worksheet558.write('R22', 'JML', body)

            worksheet558.conditional_format(22, 0, row558+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 575
            worksheet575.insert_image('A1', r'logo resmi nf.jpg')

            worksheet575.set_column('A:A', 7, center)
            worksheet575.set_column('B:B', 6, center)
            worksheet575.set_column('C:C', 18.14, center)
            worksheet575.set_column('D:D', 25, left)
            worksheet575.set_column('E:E', 13.14, left)
            worksheet575.set_column('F:F', 8.57, center)
            worksheet575.set_column('G:R', 5, center)
            worksheet575.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CINERE', title)
            worksheet575.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet575.write('A5', 'LOKASI', header)
            worksheet575.write('B5', 'TOTAL', header)
            worksheet575.merge_range('A4:B4', 'RANK', header)
            worksheet575.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet575.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet575.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet575.merge_range('F4:F5', 'KELAS', header)
            worksheet575.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet575.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet575.write('G5', 'MAT', body)
            worksheet575.write('H5', 'IND', body)
            worksheet575.write('I5', 'ENG', body)
            worksheet575.write('J5', 'IPA', body)
            worksheet575.write('K5', 'IPS', body)
            worksheet575.write('L5', 'JML', body)
            worksheet575.write('M5', 'MAT', body)
            worksheet575.write('N5', 'IND', body)
            worksheet575.write('O5', 'ENG', body)
            worksheet575.write('P5', 'IPA', body)
            worksheet575.write('Q5', 'IPS', body)
            worksheet575.write('R5', 'JML', body)

            worksheet575.conditional_format(5, 0, row575_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet575.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF CINERE', title)
            worksheet575.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet575.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet575.write('A22', 'LOKASI', header)
            worksheet575.write('B22', 'TOTAL', header)
            worksheet575.merge_range('A21:B21', 'RANK', header)
            worksheet575.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet575.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet575.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet575.merge_range('F21:F22', 'KELAS', header)
            worksheet575.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet575.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet575.write('G22', 'MAT', body)
            worksheet575.write('H22', 'IND', body)
            worksheet575.write('I22', 'ENG', body)
            worksheet575.write('J22', 'IPA', body)
            worksheet575.write('K22', 'IPS', body)
            worksheet575.write('L22', 'JML', body)
            worksheet575.write('M22', 'MAT', body)
            worksheet575.write('N22', 'IND', body)
            worksheet575.write('O22', 'ENG', body)
            worksheet575.write('P22', 'IPA', body)
            worksheet575.write('Q22', 'IPS', body)
            worksheet575.write('R22', 'JML', body)

            worksheet575.conditional_format(22, 0, row575+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 576
            worksheet576.insert_image('A1', r'logo resmi nf.jpg')

            worksheet576.set_column('A:A', 7, center)
            worksheet576.set_column('B:B', 6, center)
            worksheet576.set_column('C:C', 18.14, center)
            worksheet576.set_column('D:D', 25, left)
            worksheet576.set_column('E:E', 13.14, left)
            worksheet576.set_column('F:F', 8.57, center)
            worksheet576.set_column('G:R', 5, center)
            worksheet576.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIBINONG', title)
            worksheet576.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet576.write('A5', 'LOKASI', header)
            worksheet576.write('B5', 'TOTAL', header)
            worksheet576.merge_range('A4:B4', 'RANK', header)
            worksheet576.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet576.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet576.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet576.merge_range('F4:F5', 'KELAS', header)
            worksheet576.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet576.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet576.write('G5', 'MAT', body)
            worksheet576.write('H5', 'IND', body)
            worksheet576.write('I5', 'ENG', body)
            worksheet576.write('J5', 'IPA', body)
            worksheet576.write('K5', 'IPS', body)
            worksheet576.write('L5', 'JML', body)
            worksheet576.write('M5', 'MAT', body)
            worksheet576.write('N5', 'IND', body)
            worksheet576.write('O5', 'ENG', body)
            worksheet576.write('P5', 'IPA', body)
            worksheet576.write('Q5', 'IPS', body)
            worksheet576.write('R5', 'JML', body)

            worksheet576.conditional_format(5, 0, row576_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet576.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF CIBINONG', title)
            worksheet576.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet576.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet576.write('A22', 'LOKASI', header)
            worksheet576.write('B22', 'TOTAL', header)
            worksheet576.merge_range('A21:B21', 'RANK', header)
            worksheet576.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet576.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet576.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet576.merge_range('F21:F22', 'KELAS', header)
            worksheet576.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet576.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet576.write('G22', 'MAT', body)
            worksheet576.write('H22', 'IND', body)
            worksheet576.write('I22', 'ENG', body)
            worksheet576.write('J22', 'IPA', body)
            worksheet576.write('K22', 'IPS', body)
            worksheet576.write('L22', 'JML', body)
            worksheet576.write('M22', 'MAT', body)
            worksheet576.write('N22', 'IND', body)
            worksheet576.write('O22', 'ENG', body)
            worksheet576.write('P22', 'IPA', body)
            worksheet576.write('Q22', 'IPS', body)
            worksheet576.write('R22', 'JML', body)

            worksheet576.conditional_format(22, 0, row576+21, 17,
                                            {'type': 'no_errors', 'format': border})
            # worksheet 577
            worksheet577.insert_image('A1', r'logo resmi nf.jpg')

            worksheet577.set_column('A:A', 7, center)
            worksheet577.set_column('B:B', 6, center)
            worksheet577.set_column('C:C', 18.14, center)
            worksheet577.set_column('D:D', 25, left)
            worksheet577.set_column('E:E', 13.14, left)
            worksheet577.set_column('F:F', 8.57, center)
            worksheet577.set_column('G:R', 5, center)
            worksheet577.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF VETERAN (RUMAH SAKIT)', title)
            worksheet577.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet577.write('A5', 'LOKASI', header)
            worksheet577.write('B5', 'TOTAL', header)
            worksheet577.merge_range('A4:B4', 'RANK', header)
            worksheet577.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet577.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet577.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet577.merge_range('F4:F5', 'KELAS', header)
            worksheet577.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet577.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet577.write('G5', 'MAT', body)
            worksheet577.write('H5', 'IND', body)
            worksheet577.write('I5', 'ENG', body)
            worksheet577.write('J5', 'IPA', body)
            worksheet577.write('K5', 'IPS', body)
            worksheet577.write('L5', 'JML', body)
            worksheet577.write('M5', 'MAT', body)
            worksheet577.write('N5', 'IND', body)
            worksheet577.write('O5', 'ENG', body)
            worksheet577.write('P5', 'IPA', body)
            worksheet577.write('Q5', 'IPS', body)
            worksheet577.write('R5', 'JML', body)

            worksheet577.conditional_format(5, 0, row577_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet577.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF VETERAN (RUMAH SAKIT)', title)
            worksheet577.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet577.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet577.write('A22', 'LOKASI', header)
            worksheet577.write('B22', 'TOTAL', header)
            worksheet577.merge_range('A21:B21', 'RANK', header)
            worksheet577.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet577.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet577.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet577.merge_range('F21:F22', 'KELAS', header)
            worksheet577.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet577.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet577.write('G22', 'MAT', body)
            worksheet577.write('H22', 'IND', body)
            worksheet577.write('I22', 'ENG', body)
            worksheet577.write('J22', 'IPA', body)
            worksheet577.write('K22', 'IPS', body)
            worksheet577.write('L22', 'JML', body)
            worksheet577.write('M22', 'MAT', body)
            worksheet577.write('N22', 'IND', body)
            worksheet577.write('O22', 'ENG', body)
            worksheet577.write('P22', 'IPA', body)
            worksheet577.write('Q22', 'IPS', body)
            worksheet577.write('R22', 'JML', body)

            worksheet577.conditional_format(22, 0, row577+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 578
            worksheet578.insert_image('A1', r'logo resmi nf.jpg')

            worksheet578.set_column('A:A', 7, center)
            worksheet578.set_column('B:B', 6, center)
            worksheet578.set_column('C:C', 18.14, center)
            worksheet578.set_column('D:D', 25, left)
            worksheet578.set_column('E:E', 13.14, left)
            worksheet578.set_column('F:F', 8.57, center)
            worksheet578.set_column('G:R', 5, center)
            worksheet578.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MARTADINATA', title)
            worksheet578.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet578.write('A5', 'LOKASI', header)
            worksheet578.write('B5', 'TOTAL', header)
            worksheet578.merge_range('A4:B4', 'RANK', header)
            worksheet578.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet578.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet578.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet578.merge_range('F4:F5', 'KELAS', header)
            worksheet578.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet578.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet578.write('G5', 'MAT', body)
            worksheet578.write('H5', 'IND', body)
            worksheet578.write('I5', 'ENG', body)
            worksheet578.write('J5', 'IPA', body)
            worksheet578.write('K5', 'IPS', body)
            worksheet578.write('L5', 'JML', body)
            worksheet578.write('M5', 'MAT', body)
            worksheet578.write('N5', 'IND', body)
            worksheet578.write('O5', 'ENG', body)
            worksheet578.write('P5', 'IPA', body)
            worksheet578.write('Q5', 'IPS', body)
            worksheet578.write('R5', 'JML', body)

            worksheet578.conditional_format(5, 0, row578_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet578.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF MARTADINATA', title)
            worksheet578.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet578.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet578.write('A22', 'LOKASI', header)
            worksheet578.write('B22', 'TOTAL', header)
            worksheet578.merge_range('A21:B21', 'RANK', header)
            worksheet578.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet578.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet578.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet578.merge_range('F21:F22', 'KELAS', header)
            worksheet578.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet578.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet578.write('G22', 'MAT', body)
            worksheet578.write('H22', 'IND', body)
            worksheet578.write('I22', 'ENG', body)
            worksheet578.write('J22', 'IPA', body)
            worksheet578.write('K22', 'IPS', body)
            worksheet578.write('L22', 'JML', body)
            worksheet578.write('M22', 'MAT', body)
            worksheet578.write('N22', 'IND', body)
            worksheet578.write('O22', 'ENG', body)
            worksheet578.write('P22', 'IPA', body)
            worksheet578.write('Q22', 'IPS', body)
            worksheet578.write('R22', 'JML', body)

            worksheet578.conditional_format(22, 0, row578+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 588
            worksheet588.insert_image('A1', r'logo resmi nf.jpg')

            worksheet588.set_column('A:A', 7, center)
            worksheet588.set_column('B:B', 6, center)
            worksheet588.set_column('C:C', 18.14, center)
            worksheet588.set_column('D:D', 25, left)
            worksheet588.set_column('E:E', 13.14, left)
            worksheet588.set_column('F:F', 8.57, center)
            worksheet588.set_column('G:R', 5, center)
            worksheet588.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MAHARAJA', title)
            worksheet588.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet588.write('A5', 'LOKASI', header)
            worksheet588.write('B5', 'TOTAL', header)
            worksheet588.merge_range('A4:B4', 'RANK', header)
            worksheet588.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet588.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet588.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet588.merge_range('F4:F5', 'KELAS', header)
            worksheet588.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet588.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet588.write('G5', 'MAT', body)
            worksheet588.write('H5', 'IND', body)
            worksheet588.write('I5', 'ENG', body)
            worksheet588.write('J5', 'IPA', body)
            worksheet588.write('K5', 'IPS', body)
            worksheet588.write('L5', 'JML', body)
            worksheet588.write('M5', 'MAT', body)
            worksheet588.write('N5', 'IND', body)
            worksheet588.write('O5', 'ENG', body)
            worksheet588.write('P5', 'IPA', body)
            worksheet588.write('Q5', 'IPS', body)
            worksheet588.write('R5', 'JML', body)

            worksheet588.conditional_format(5, 0, row588_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet588.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF MAHARAJA', title)
            worksheet588.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet588.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet588.write('A22', 'LOKASI', header)
            worksheet588.write('B22', 'TOTAL', header)
            worksheet588.merge_range('A21:B21', 'RANK', header)
            worksheet588.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet588.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet588.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet588.merge_range('F21:F22', 'KELAS', header)
            worksheet588.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet588.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet588.write('G22', 'MAT', body)
            worksheet588.write('H22', 'IND', body)
            worksheet588.write('I22', 'ENG', body)
            worksheet588.write('J22', 'IPA', body)
            worksheet588.write('K22', 'IPS', body)
            worksheet588.write('L22', 'JML', body)
            worksheet588.write('M22', 'MAT', body)
            worksheet588.write('N22', 'IND', body)
            worksheet588.write('O22', 'ENG', body)
            worksheet588.write('P22', 'IPA', body)
            worksheet588.write('Q22', 'IPS', body)
            worksheet588.write('R22', 'JML', body)

            worksheet588.conditional_format(22, 0, row588+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 589
            worksheet589.insert_image('A1', r'logo resmi nf.jpg')

            worksheet589.set_column('A:A', 7, center)
            worksheet589.set_column('B:B', 6, center)
            worksheet589.set_column('C:C', 18.14, center)
            worksheet589.set_column('D:D', 25, left)
            worksheet589.set_column('E:E', 13.14, left)
            worksheet589.set_column('F:F', 8.57, center)
            worksheet589.set_column('G:R', 5, center)
            worksheet589.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF WARUNG JAMBU', title)
            worksheet589.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet589.write('A5', 'LOKASI', header)
            worksheet589.write('B5', 'TOTAL', header)
            worksheet589.merge_range('A4:B4', 'RANK', header)
            worksheet589.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet589.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet589.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet589.merge_range('F4:F5', 'KELAS', header)
            worksheet589.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet589.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet589.write('G5', 'MAT', body)
            worksheet589.write('H5', 'IND', body)
            worksheet589.write('I5', 'ENG', body)
            worksheet589.write('J5', 'IPA', body)
            worksheet589.write('K5', 'IPS', body)
            worksheet589.write('L5', 'JML', body)
            worksheet589.write('M5', 'MAT', body)
            worksheet589.write('N5', 'IND', body)
            worksheet589.write('O5', 'ENG', body)
            worksheet589.write('P5', 'IPA', body)
            worksheet589.write('Q5', 'IPS', body)
            worksheet589.write('R5', 'JML', body)

            worksheet589.conditional_format(5, 0, row589_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet589.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF WARUNG JAMBU', title)
            worksheet589.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet589.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet589.write('A22', 'LOKASI', header)
            worksheet589.write('B22', 'TOTAL', header)
            worksheet589.merge_range('A21:B21', 'RANK', header)
            worksheet589.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet589.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet589.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet589.merge_range('F21:F22', 'KELAS', header)
            worksheet589.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet589.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet589.write('G22', 'MAT', body)
            worksheet589.write('H22', 'IND', body)
            worksheet589.write('I22', 'ENG', body)
            worksheet589.write('J22', 'IPA', body)
            worksheet589.write('K22', 'IPS', body)
            worksheet589.write('L22', 'JML', body)
            worksheet589.write('M22', 'MAT', body)
            worksheet589.write('N22', 'IND', body)
            worksheet589.write('O22', 'ENG', body)
            worksheet589.write('P22', 'IPA', body)
            worksheet589.write('Q22', 'IPS', body)
            worksheet589.write('R22', 'JML', body)

            worksheet589.conditional_format(22, 0, row589+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 594
            worksheet594.insert_image('A1', r'logo resmi nf.jpg')

            worksheet594.set_column('A:A', 7, center)
            worksheet594.set_column('B:B', 6, center)
            worksheet594.set_column('C:C', 18.14, center)
            worksheet594.set_column('D:D', 25, left)
            worksheet594.set_column('E:E', 13.14, left)
            worksheet594.set_column('F:F', 8.57, center)
            worksheet594.set_column('G:R', 5, center)
            worksheet594.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PEMDA CIBINONG', title)
            worksheet594.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet594.write('A5', 'LOKASI', header)
            worksheet594.write('B5', 'TOTAL', header)
            worksheet594.merge_range('A4:B4', 'RANK', header)
            worksheet594.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet594.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet594.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet594.merge_range('F4:F5', 'KELAS', header)
            worksheet594.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet594.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet594.write('G5', 'MAT', body)
            worksheet594.write('H5', 'IND', body)
            worksheet594.write('I5', 'ENG', body)
            worksheet594.write('J5', 'IPA', body)
            worksheet594.write('K5', 'IPS', body)
            worksheet594.write('L5', 'JML', body)
            worksheet594.write('M5', 'MAT', body)
            worksheet594.write('N5', 'IND', body)
            worksheet594.write('O5', 'ENG', body)
            worksheet594.write('P5', 'IPA', body)
            worksheet594.write('Q5', 'IPS', body)
            worksheet594.write('R5', 'JML', body)

            worksheet594.conditional_format(5, 0, row594_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet594.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF PEMDA CIBINONG', title)
            worksheet594.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet594.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet594.write('A22', 'LOKASI', header)
            worksheet594.write('B22', 'TOTAL', header)
            worksheet594.merge_range('A21:B21', 'RANK', header)
            worksheet594.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet594.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet594.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet594.merge_range('F21:F22', 'KELAS', header)
            worksheet594.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet594.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet594.write('G22', 'MAT', body)
            worksheet594.write('H22', 'IND', body)
            worksheet594.write('I22', 'ENG', body)
            worksheet594.write('J22', 'IPA', body)
            worksheet594.write('K22', 'IPS', body)
            worksheet594.write('L22', 'JML', body)
            worksheet594.write('M22', 'MAT', body)
            worksheet594.write('N22', 'IND', body)
            worksheet594.write('O22', 'ENG', body)
            worksheet594.write('P22', 'IPA', body)
            worksheet594.write('Q22', 'IPS', body)
            worksheet594.write('R22', 'JML', body)

            worksheet594.conditional_format(22, 0, row594+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 661
            worksheet661.insert_image('A1', r'logo resmi nf.jpg')

            worksheet661.set_column('A:A', 7, center)
            worksheet661.set_column('B:B', 6, center)
            worksheet661.set_column('C:C', 18.14, center)
            worksheet661.set_column('D:D', 25, left)
            worksheet661.set_column('E:E', 13.14, left)
            worksheet661.set_column('F:F', 8.57, center)
            worksheet661.set_column('G:R', 5, center)
            worksheet661.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF GAJAH MADA', title)
            worksheet661.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet661.write('A5', 'LOKASI', header)
            worksheet661.write('B5', 'TOTAL', header)
            worksheet661.merge_range('A4:B4', 'RANK', header)
            worksheet661.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet661.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet661.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet661.merge_range('F4:F5', 'KELAS', header)
            worksheet661.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet661.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet661.write('G5', 'MAT', body)
            worksheet661.write('H5', 'IND', body)
            worksheet661.write('I5', 'ENG', body)
            worksheet661.write('J5', 'IPA', body)
            worksheet661.write('K5', 'IPS', body)
            worksheet661.write('L5', 'JML', body)
            worksheet661.write('M5', 'MAT', body)
            worksheet661.write('N5', 'IND', body)
            worksheet661.write('O5', 'ENG', body)
            worksheet661.write('P5', 'IPA', body)
            worksheet661.write('Q5', 'IPS', body)
            worksheet661.write('R5', 'JML', body)

            worksheet661.conditional_format(5, 0, row661_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet661.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF GAJAH MADA', title)
            worksheet661.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet661.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet661.write('A22', 'LOKASI', header)
            worksheet661.write('B22', 'TOTAL', header)
            worksheet661.merge_range('A21:B21', 'RANK', header)
            worksheet661.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet661.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet661.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet661.merge_range('F21:F22', 'KELAS', header)
            worksheet661.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet661.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet661.write('G22', 'MAT', body)
            worksheet661.write('H22', 'IND', body)
            worksheet661.write('I22', 'ENG', body)
            worksheet661.write('J22', 'IPA', body)
            worksheet661.write('K22', 'IPS', body)
            worksheet661.write('L22', 'JML', body)
            worksheet661.write('M22', 'MAT', body)
            worksheet661.write('N22', 'IND', body)
            worksheet661.write('O22', 'ENG', body)
            worksheet661.write('P22', 'IPA', body)
            worksheet661.write('Q22', 'IPS', body)
            worksheet661.write('R22', 'JML', body)

            worksheet661.conditional_format(22, 0, row661+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 662
            worksheet662.insert_image('A1', r'logo resmi nf.jpg')

            worksheet662.set_column('A:A', 7, center)
            worksheet662.set_column('B:B', 6, center)
            worksheet662.set_column('C:C', 18.14, center)
            worksheet662.set_column('D:D', 25, left)
            worksheet662.set_column('E:E', 13.14, left)
            worksheet662.set_column('F:F', 8.57, center)
            worksheet662.set_column('G:R', 5, center)
            worksheet662.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF LOLONG BELANTI', title)
            worksheet662.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet662.write('A5', 'LOKASI', header)
            worksheet662.write('B5', 'TOTAL', header)
            worksheet662.merge_range('A4:B4', 'RANK', header)
            worksheet662.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet662.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet662.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet662.merge_range('F4:F5', 'KELAS', header)
            worksheet662.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet662.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet662.write('G5', 'MAT', body)
            worksheet662.write('H5', 'IND', body)
            worksheet662.write('I5', 'ENG', body)
            worksheet662.write('J5', 'IPA', body)
            worksheet662.write('K5', 'IPS', body)
            worksheet662.write('L5', 'JML', body)
            worksheet662.write('M5', 'MAT', body)
            worksheet662.write('N5', 'IND', body)
            worksheet662.write('O5', 'ENG', body)
            worksheet662.write('P5', 'IPA', body)
            worksheet662.write('Q5', 'IPS', body)
            worksheet662.write('R5', 'JML', body)

            worksheet662.conditional_format(5, 0, row662_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet662.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF LOLONG BELANTI', title)
            worksheet662.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet662.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet662.write('A22', 'LOKASI', header)
            worksheet662.write('B22', 'TOTAL', header)
            worksheet662.merge_range('A21:B21', 'RANK', header)
            worksheet662.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet662.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet662.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet662.merge_range('F21:F22', 'KELAS', header)
            worksheet662.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet662.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet662.write('G22', 'MAT', body)
            worksheet662.write('H22', 'IND', body)
            worksheet662.write('I22', 'ENG', body)
            worksheet662.write('J22', 'IPA', body)
            worksheet662.write('K22', 'IPS', body)
            worksheet662.write('L22', 'JML', body)
            worksheet662.write('M22', 'MAT', body)
            worksheet662.write('N22', 'IND', body)
            worksheet662.write('O22', 'ENG', body)
            worksheet662.write('P22', 'IPA', body)
            worksheet662.write('Q22', 'IPS', body)
            worksheet662.write('R22', 'JML', body)

            worksheet662.conditional_format(22, 0, row662+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 663
            worksheet663.insert_image('A1', r'logo resmi nf.jpg')

            worksheet663.set_column('A:A', 7, center)
            worksheet663.set_column('B:B', 6, center)
            worksheet663.set_column('C:C', 18.14, center)
            worksheet663.set_column('D:D', 25, left)
            worksheet663.set_column('E:E', 13.14, left)
            worksheet663.set_column('F:F', 8.57, center)
            worksheet663.set_column('G:R', 5, center)
            worksheet663.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SOETOMO', title)
            worksheet663.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet663.write('A5', 'LOKASI', header)
            worksheet663.write('B5', 'TOTAL', header)
            worksheet663.merge_range('A4:B4', 'RANK', header)
            worksheet663.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet663.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet663.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet663.merge_range('F4:F5', 'KELAS', header)
            worksheet663.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet663.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet663.write('G5', 'MAT', body)
            worksheet663.write('H5', 'IND', body)
            worksheet663.write('I5', 'ENG', body)
            worksheet663.write('J5', 'IPA', body)
            worksheet663.write('K5', 'IPS', body)
            worksheet663.write('L5', 'JML', body)
            worksheet663.write('M5', 'MAT', body)
            worksheet663.write('N5', 'IND', body)
            worksheet663.write('O5', 'ENG', body)
            worksheet663.write('P5', 'IPA', body)
            worksheet663.write('Q5', 'IPS', body)
            worksheet663.write('R5', 'JML', body)

            worksheet663.conditional_format(5, 0, row663_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet663.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF SOETOMO', title)
            worksheet663.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet663.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet663.write('A22', 'LOKASI', header)
            worksheet663.write('B22', 'TOTAL', header)
            worksheet663.merge_range('A21:B21', 'RANK', header)
            worksheet663.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet663.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet663.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet663.merge_range('F21:F22', 'KELAS', header)
            worksheet663.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet663.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet663.write('G22', 'MAT', body)
            worksheet663.write('H22', 'IND', body)
            worksheet663.write('I22', 'ENG', body)
            worksheet663.write('J22', 'IPA', body)
            worksheet663.write('K22', 'IPS', body)
            worksheet663.write('L22', 'JML', body)
            worksheet663.write('M22', 'MAT', body)
            worksheet663.write('N22', 'IND', body)
            worksheet663.write('O22', 'ENG', body)
            worksheet663.write('P22', 'IPA', body)
            worksheet663.write('Q22', 'IPS', body)
            worksheet663.write('R22', 'JML', body)

            worksheet663.conditional_format(22, 0, row663+21, 17,
                                            {'type': 'no_errors', 'format': border})

            # worksheet 664
            worksheet664.insert_image('A1', r'logo resmi nf.jpg')

            worksheet664.set_column('A:A', 7, center)
            worksheet664.set_column('B:B', 6, center)
            worksheet664.set_column('C:C', 18.14, center)
            worksheet664.set_column('D:D', 25, left)
            worksheet664.set_column('E:E', 13.14, left)
            worksheet664.set_column('F:F', 8.57, center)
            worksheet664.set_column('G:R', 5, center)
            worksheet664.merge_range(
                'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF TAN MALAKA', title)
            worksheet664.merge_range(
                'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
            worksheet664.write('A5', 'LOKASI', header)
            worksheet664.write('B5', 'TOTAL', header)
            worksheet664.merge_range('A4:B4', 'RANK', header)
            worksheet664.merge_range('C4:C5', 'NOMOR NF', header)
            worksheet664.merge_range('D4:D5', 'NAMA SISWA', header)
            worksheet664.merge_range('E4:E5', 'SEKOLAH', header)
            worksheet664.merge_range('F4:F5', 'KELAS', header)
            worksheet664.merge_range('G4:L4', 'JUMLAH BENAR', header)
            worksheet664.merge_range('M4:R4', 'NILAI STANDAR', header)
            worksheet664.write('G5', 'MAT', body)
            worksheet664.write('H5', 'IND', body)
            worksheet664.write('I5', 'ENG', body)
            worksheet664.write('J5', 'IPA', body)
            worksheet664.write('K5', 'IPS', body)
            worksheet664.write('L5', 'JML', body)
            worksheet664.write('M5', 'MAT', body)
            worksheet664.write('N5', 'IND', body)
            worksheet664.write('O5', 'ENG', body)
            worksheet664.write('P5', 'IPA', body)
            worksheet664.write('Q5', 'IPS', body)
            worksheet664.write('R5', 'JML', body)

            worksheet664.conditional_format(5, 0, row664_10+4, 17,
                                            {'type': 'no_errors', 'format': border})

            worksheet664.merge_range(
                'A17:R17', fr'KELAS {kelas} - LOKASI NF TAN MALAKA', title)
            worksheet664.merge_range('A18:R18', fr'{penilaian}', subTitle)
            worksheet664.merge_range(
                'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
            worksheet664.write('A22', 'LOKASI', header)
            worksheet664.write('B22', 'TOTAL', header)
            worksheet664.merge_range('A21:B21', 'RANK', header)
            worksheet664.merge_range('C21:C22', 'NOMOR NF', header)
            worksheet664.merge_range('D21:D22', 'NAMA SISWA', header)
            worksheet664.merge_range('E21:E22', 'SEKOLAH', header)
            worksheet664.merge_range('F21:F22', 'KELAS', header)
            worksheet664.merge_range('G21:L21', 'JUMLAH BENAR', header)
            worksheet664.merge_range('M21:R21', 'NILAI STANDAR', header)
            worksheet664.write('G22', 'MAT', body)
            worksheet664.write('H22', 'IND', body)
            worksheet664.write('I22', 'ENG', body)
            worksheet664.write('J22', 'IPA', body)
            worksheet664.write('K22', 'IPS', body)
            worksheet664.write('L22', 'JML', body)
            worksheet664.write('M22', 'MAT', body)
            worksheet664.write('N22', 'IND', body)
            worksheet664.write('O22', 'ENG', body)
            worksheet664.write('P22', 'IPA', body)
            worksheet664.write('Q22', 'IPS', body)
            worksheet664.write('R22', 'JML', body)

            worksheet664.conditional_format(22, 0, row664+21, 17,
                                            {'type': 'no_errors', 'format': border})

            workbook.close()
            st.success("File siap diunduh!")

            # Tombol unduh file
            with open(file_path, "rb") as f:
                bytes_data = f.read()
            st.download_button(label="Unduh File", data=bytes_data,
                               file_name=file_name)