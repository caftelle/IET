###########################################################################
########################## DEVELOPED BY CAFTELLE ##########################
########################## DEVELOPED BY CAFTELLE ##########################
########################## DEVELOPED BY CAFTELLE ##########################
###########################################################################
################################# I E T ###################################
######################## V E R S I O N  1 . 0 . 5 #########################
###########################################################################
###########################################################################


# GEREKLİ KÜTÜPHANELER


import cv2 as cv
from pyzbar import pyzbar
from pyzbar.pyzbar import decode
import pytesseract
import xlsxwriter
import openpyxl
import datetime
import requests
import os
import mimetypes
from email.message import EmailMessage
import smtplib
import beepy as beep


print('')
print('#############################################################################################################################')
print('################################################### DEVELOPED BY CAFTELLE ###################################################')
print('#############################################################################################################################')
print('')
def dosyasilme():
  try:
    print(' | Taranan Dosyayı silmek bir önceki ve sonraki taramalardaki verilerin karışmaması için önemlidir. Silinmesi önerilir.')
    print(' | [ 5 ] Taranan dosyayı silmek istiyorum. | [ 6 ] Taranan dosyayı silmek istemiyorum.')
    sec3 = input(' | Lütfen bir seçim yapınız : ')

    if sec3 == '5':

        tarananisemri2 = 'Taranan_Is_Emirleri.xlsx'
        os.remove(tarananisemri2)
        print(' | Taranan_Is_Emirleri dosyasını sildim. Çıkış yapıyorum.')

    elif sec3 == '6':
        print(' | Dosyayı silmeyeceğim. Çıkış yapıyorum.')
    else:

        print(' | Hatalı bir kod tuşladınız tekrar deneyiniz.')
        beep.beep(3)
        dosyasilme()
  except:

      print(' | Dosya zaten silinmiş.')

      return False
def MailGonder():

    try:
        
        print(' | [ 3 ] Mail adresini yazmak istiyorum. | [ 4 ] Kullanıcı Adımı yazmak istiyorum.')
        sec = input(' | Lütfen bir seçim yapınız : ')
        sec = str(sec)

        if sec == '3':
            sec2 = input(' | Lütfen Mail Adresi giriniz: ')
            mail = sec2
        elif sec == '4':
            sec = input(' | Kullanıcı Adınızı Giriniz : ')
            mail = sec + '@turksat.com.tr'
        else:

            print(' | Hatalı kod tuşladınız. Tekrar Deneyiniz.')
            beep.beep(3)
            MailGonder()

        recipient = mail
        bdtarih = datetime.datetime.now()
        yil = bdtarih.year
        ay = bdtarih.month
        gun = bdtarih.day
        saat = bdtarih.hour
        dakika = bdtarih.minute
        toplami = str(yil) + '_' + str(ay) + '_' + str(gun)
        dosyaadi = toplami + '_Tarihli_Is_Emirleri_Tutanagı.xlsm'
        dosyaadifinal = str(dosyaadi)

        # Yazılan Dosyayı Arama
        tutanakdizinpath2 = str(os.getcwd())
        tutanakdosyasi2 = tutanakdizinpath2 + '/' + 'TutanakForm.xlsm'

        if os.path.isfile(tutanakdosyasi2):

            tttarih = datetime.datetime.now()
            ttyil = tttarih.year
            ttay = tttarih.month
            ttgun = tttarih.day
            ttsaat = tttarih.hour
            ttdakika = tttarih.minute
            tttoplami = str(ttyil) + '_' + str(ttay) + '_' + str(ttgun)
            ttdosyaadi = tttoplami + '_Tarihli_Is_Emirleri_Tutanagı.xlsm'
            ttdosyaadifinal = str(ttdosyaadi)
            tarananisemri = 'Taranan_Is_Emirleri.xlsx'

            print(' | Gönderilecek Dosya Bulundu. Mail göndermeye hazırlanıyorum. ')
            mail_server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
            mail_server.login("developed.by.caftelle@gmail.com", 'yzgchfnfzbivbhei')
            message = EmailMessage()
            sender = "developed.by.caftelle@gmail.com"
            recipient = mail
            message['From'] = 'Caftelle Software'
            message['To'] = recipient
            message['Subject'] = toplami + ' Tarihli İş Emirleri Tutanağı'
            body = 'Merhabalar, \n\n' + toplami + ' Tarihli İş Emirleri Ektedir.\n\nİyi Calismalar. \n \n \n | Developed by Caftelle  \n | Caftelle Created by Furkan ARINCI'
            message.set_content(body)
            mime_type, _ = mimetypes.guess_type(ttdosyaadifinal)
            mime_type, mime_subtype = mime_type.split('/')

            with open(tutanakdosyasi2, 'rb') as file:
                message.add_attachment(file.read(), maintype=mime_type, subtype=mime_subtype,
                                       filename=tttoplami + '_Tarihli_Is_Emirleri_Tutanagı.xlsm')
                print(' | Taranan Tutanak Formunu maile ekledim... ')

            with open(tarananisemri, 'rb') as file:
                print(' | Taranan Is Emirleri Formunu maile ekledim... ')
                message.add_attachment(file.read(), maintype=mime_type, subtype=mime_subtype,
                                           filename='Taranan_Is_Emirleri.xlsx')

            mail_server.send_message(message)
            mail_server.quit()

            print('\n | Gönderilen Mail Adresi: '+ recipient +'\n | M A I L  B A S A R I Y L A  G O N D E R I L D I . \n')
            beep.beep(5)


        else:

            print(' | Dosya bulunamadığı için mail gönderilemedi.')
            beep.beep(3)

    except:

        print(' | Mail Adresini veya Kullanıcı Adı yanlış olduğu için mail gönderilemedi.')
        beep.beep(3)
        MailGonder()

def AllWithLove():


    print('')
    print(
        '################################################### DEVELOPED BY CAFTELLE ###################################################')
    print('')
    # FormDosyaAdıBelirleme
    bdtarih = datetime.datetime.now()
    yil = bdtarih.year
    ay = bdtarih.month
    gun = bdtarih.day
    saat = bdtarih.hour
    dakika = bdtarih.minute
    toplami = str(yil) + '_' + str(ay) + '_' + str(gun)
    dosyaadi = toplami + '_Tarihli_Is_Emirleri_Tutanagı.xlsx'
    dosyaadifinal = str(dosyaadi)

    tutanakdizinpath = str(os.getcwd())
    tutanakdosyasi = tutanakdizinpath + '/' + 'TutanakForm.xlsm'

    for root, dir, files in os.walk(tutanakdizinpath):

        if 'TutanakForm.xlsm' in files:
            print(' | Tutanak Formu Dosya içerisinde mevcut. İşleme devam ediyorum... ')
            break

        if not 'TutanakForm.xlsm' in files:
            print('| Tutanak Formu bulunamadı. Hemen İndiriyorum... ')
            # FormDosyasıİndirme
            print('| Tutanak Formu indiriliyor... ')
            resp = requests.get(
                'https://www.dropbox.com/scl/fi/ydyx5isxb2szdhntamf85/TutanakForm.xlsm?dl=1&rlkey=g63u0w9uago9jajslnpejkjqd')

            with open('TutanakForm.xlsm', 'wb') as output:
                output.write(resp.content)

                print('| İndirme Tamamlandı. ')
                beep.beep(5)

            break



    # Excel Satır Döngüsü
    iptalline = 2
    tesisline = 2


    # Excel Dosyası Oluşturma
    planWorkbook1 = xlsxwriter.Workbook('Taranan_Is_Emirleri.xlsx')
    planSheettesis12 = planWorkbook1.add_worksheet("TESIS")
    planSheetiptal12 = planWorkbook1.add_worksheet("IPTAL")
    planWorkbook1.close()
    planWorkbook = openpyxl.load_workbook('Taranan_Is_Emirleri.xlsx')

    # Excel Stün ve Sekme Oluşturma
    planSheettesis = planWorkbook["TESIS"]
    planSheetiptal = planWorkbook["IPTAL"]

    planSheettesis['A1'] = 'Hizmet Numarası'
    planSheettesis['B1'] = 'Müşteri Numarası'
    planSheettesis['C1'] = 'İş Emri Numarası'
    planSheettesis['C1'] = 'İş Emri Numarası'
    planSheettesis['D1'] = 'Hizmet Türü'
    planSheettesis['E1'] = 'İş Emri Tipi'
    planSheettesis['F1'] = 'Tarih'

    planSheetiptal['A1'] = 'Hizmet Numarası'
    planSheetiptal['B1'] = 'Müşteri Numarası'
    planSheetiptal['C1'] = 'İş Emri Numarası'
    planSheetiptal['D1'] = 'Hizmet Türü'
    planSheetiptal['E1'] = 'İş Emri Tipi'
    planSheetiptal['F1'] = 'Tarih'

    # QR TARAMA


    while True:

        cap1 = cv.VideoCapture(0)
        cap1.set(3, 1920)
        cap1.set(4, 1080)

        # Değerlerin Sıfırlanması
        textstart = False
        savestart = True
        qrstart = True
        gerekli =False
        musterinofinal = '(     )'
        hizmetnofinal = '(     )'
        isemrinofinal = '1'
        isemrituru = '1'
        iptalturu = '(     )'
        isemriturufinal = '(     )'
        musterinoindex = 0
        isemrinoindex = 0
        qrhizmetno = 0
        isemrinoindex = 0
        isemrituruindex = 0


        while qrstart:

            # Değerleri Sıfırlama
            musterinoindex = 0
            isemrinoindex = 0
            qrhizmetno = 0
            isemrinoindex = 0
            isemrituruindex = 0

            # Kamera'dan Aldığı Veriyi Okuma

            success, img = cap1.read()

            font = cv.FONT_ITALIC
            decodedObjects = pyzbar.decode(img)

            for obj in decodedObjects:
                qrtemiz2 = obj.data.decode('utf-8')
                cv.putText(img, str(qrtemiz2), (200, 200), font, 1,
                           (0, 255, 160), 2)
                beep.beep(1)


            print(
                ' | Müşteri No Taranıyor... ' + ' | Hizmet No Taranıyor... ' + '| İş Emri No Taranıyor... ' + ' | İş Emri Türü bir sonraki aşamada taranacak.')
            cv.imshow("QR Tarama", img)
            cv.waitKey(1)

            for qrcodee in decode(img):

                # Kamera'dan Alınan Verideki Yazıları Okuma
                print(' | QR Okundu ve Analiz Ediliyor...')
                qrtemiz = qrcodee.data.decode('utf-8')
                qrlist = qrtemiz.split("|")
                qrlistno = len(qrlist)

                # MÜSTERİ NO AYIKLAMA QR
                musterinoindex = [datano for datano in range(qrlistno) if qrlist[datano].startswith('M')]
                qrmusterino = qrlist[musterinoindex[0]]
                musterinofinal = qrmusterino.replace("M", "")
                print(' | Taranan QR Code içinden Müşteri Numarası ayıklanıyor...')

                # HİZMET NO AYIKLAMA QR
                hizmetnoindex = [datano for datano in range(qrlistno) if qrlist[datano].startswith('H')]
                qrhizmetno = qrlist[hizmetnoindex[0]]
                hizmetnofinal = qrhizmetno.replace("H", "")
                print(' | Taranan QR Code içinden Hizmet Numarası ayıklanıyor...')

                try:
                    # İş Emri Türü AYIKLAMA QR Eğer Eklenirse
                    isemrituruindex = [datano for datano in range(qrlistno) if qrlist[datano].startswith('IT')]
                    isemrituru1 = qrlist[isemrituruindex[0]]
                    isemriturufinal = isemrituru1.replace("IT", "")

                    replace_chars = [('ı', 'i'), ('İ', 'I'), ('ü', 'u'), ('Ü', 'U'), ('ö', 'o'), ('Ö', 'O'), ('ç', 'c'),
                                         ('Ç', 'C'),
                                         ('ş', 's'), ('Ş', 'S'), ('ğ', 'g'), ('Ğ', 'G')]

                    for search, replace in replace_chars:

                            isemriturufinal = isemriturufinal.replace(search, replace)
                            isemriturufinal = isemriturufinal
                            isemriturufinal = isemriturufinal.upper()
                            isemriturufinal = isemriturufinal.strip()
                            text2 = isemriturufinal
                except:
                    pass

                # İŞ EMRİ NO AYIKLAMA QR
                isemrinoindex = [datano for datano in range(qrlistno) if qrlist[datano].startswith('I')]
                isemrinoindex2 = len(isemrinoindex)
                if isemrinoindex2 > 0:
                    qrisemrino = qrlist[isemrinoindex[0]]
                    isemrinofinal = qrisemrino.replace("I", "")
                    print(' | Taranan QR Code içinden İş Emri Numarası ayıklanıyor...')

                # Müşteri Numarasının ve Hizmet Numarasının Alındığını Doğrulama
                if musterinofinal != '(     )' and hizmetnofinal != '(     )':
                    qrstart = False

        if musterinofinal != '(     )' and hizmetnofinal != '(     )':
            cv.destroyAllWindows()
            textstart = True

        if isemriturufinal != '(     )':
            textstart = False
            gerekli = True

        # YAZI TARAMA

        while textstart:

            # Kamera'dan Yazıları Okuma
            success, img = cap1.read()
            h, w, _ = img.shape
            print(
                '| Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: Taranıyor... Lütfen kameraya gösteriniz.')
            text2 = pytesseract.image_to_string(img)
            boxes = pytesseract.image_to_boxes(img)

            for b in boxes.splitlines():
                b = b.split(' ')
                img = cv.rectangle(img, (int(b[1]), h - int(b[2])), (int(b[3]), h - int(b[4])), (0, 255, 160), 1)

            replace_chars = [('ı', 'i'), ('İ', 'I'), ('ü', 'u'), ('Ü', 'U'), ('ö', 'o'), ('Ö', 'O'), ('ç', 'c'),
                             ('Ç', 'C'),
                             ('ş', 's'), ('Ş', 'S'), ('ğ', 'g'), ('Ğ', 'G')]
            for search, replace in replace_chars:
                text2 = text2.replace(search, replace)
                text2 = text2.upper()
                text2 = text2.strip()

                break

            cv.imshow("Yazi Tarama", img)
            cv.waitKey(1)

            # Okunan Yazıları Tanıma ve Türüne Göre Ayıklama
            for img in text2:

                if 'OKUNAMIYOR' in text2:

                    iptalturu = 'OKUNAMADI'
                    textstart = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + 'OKUNMUYOR.')
                    break

                if 'NUMARA TASIMA' in text2:
                    iptalturu = 'NUMARA TAŞIMALI YENİ ABONELİK'
                    textstart = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break
                if 'KABLOSES IPTAL' in text2:
                    iptalturu = 'KABLOSES İPTAL'
                    textstart = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break

                if 'ABONE ISTEGI ILE KABLOSES IPTALI' in text2:
                    iptalturu = 'KABLOSES İPTAL'
                    textstart = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break

                if 'VERASETEN' and 'VERASETEN IPTAL' and 'VERASET' in text2:
                    iptalturu = 'VERASETEN İPTAL'
                    textstart = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break

                if  'ABONE ISTEGI ILE IPTAL' in text2:
                    iptalturu = 'ABONELİK İPTAL'
                    textstart = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break

                if  'ABONELIK IPTAL' in  text2:
                    iptalturu = 'ABONELİK İPTAL'
                    textstart = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break

                if  'TARIFE DEGISIMI' in text2:
                    iptalturu = 'TARİFE DEĞİŞİMİ'
                    textstart = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break
                if  'KAMPANYAYA GECIS' in text2:
                    iptalturu = 'TARİFE DEĞİŞİMİ'
                    textstart = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break
                if 'KIRALAMA IPTAL' and 'CIHAZ KIRALAMA IPTAL'in text2:
                    iptalturu = 'CİHAZ KİRALAMA İPTAL'
                    textstart = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break
                if 'KIRALAMA SIPARIS' and 'CIHAZ KIRALAMA SIPARIS' in text2:
                    iptalturu = 'CİHAZ KİRALAMA SİPARİŞ'
                    textstart = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break
                if 'TAAHHUTLU ABONELIK DEVIR ALMA' and 'ABONELIK DEVIR ALMA' in text2:
                    iptalturu = 'TAAHHÜTLÜ ABONELİK DEVİR ALMA'
                    textstart = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break
                if 'TAAHHUTLU ABONELIK DEVIR' and 'ABONELIK DEVIR' in text2:
                    iptalturu = 'TAAHHÜTLÜ ABONELİK DEVİR'
                    textstart = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break
                if  'YENI ABONELIK' in text2:
                    iptalturu = 'TAAHHÜTLÜ YENİ ABONELİK'
                    textstart = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        '| Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break
                if 'NAKIL GELEN'in text2:
                    iptalturu = 'TAAHHÜTLÜ NAKİL GELEN'
                    textstart = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break
                if 'ASKIYA ALMA' in text2:
                    iptalturu = 'HİZMETİ ASKIYA ALMA'
                    textstart = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break
                if 'HIZMETI ASKIYA ALMA' in text2:
                    iptalturu = 'HİZMETİ ASKIYA ALMA'
                    textstart = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break
                if 'CIHAZ IADE' and 'CIHAZ IADE FORMU' and 'IADE FORMU' in text2:
                    iptalturu = 'CİHAZ İADE FORMU'
                    textstart = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break
                if 'KABLONET IPTAL BAŞVURU FORMU' and 'KABLOTV IPTAL BASVURU FORMU' and 'IPTAL BASVURU' in text2:
                    iptalturu = 'KABLONET İPTAL FORMU'
                    textstart = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break

        while gerekli:
            print('gerekli')
            replace_chars = [('ı', 'i'), ('İ', 'I'), ('ü', 'u'), ('Ü', 'U'), ('ö', 'o'), ('Ö', 'O'), ('ç', 'c'),
                             ('Ç', 'C'),
                             ('ş', 's'), ('Ş', 'S'), ('ğ', 'g'), ('Ğ', 'G')]
            for search, replace in replace_chars:
                text2 = text2.replace(search, replace)
                text2 = text2
                text2 = text2.upper()
                text2 = text2.strip()
                break

            # Okunan Yazıları Tanıma ve Türüne Göre Ayıklama
            for img in text2:

                if 'OKUNAMIYOR' in text2:

                    iptalturu = 'OKUNAMADI'
                    gerekli = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + 'OKUNMUYOR.')
                    break

                if 'NUMARATASIMA' in text2:

                    iptalturu = 'NUMARA TAŞIMALI YENİ ABONELİK'
                    gerekli = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break

                if 'KABLOSESIPTAL' in text2:
                    iptalturu = 'KABLOSES İPTAL'
                    textstart = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break

                if  'VERASETENIPTAL' in text2:
                    iptalturu = 'VERASETEN İPTAL'
                    gerekli = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break
                if 'ABONELIKIPTAL' in text2:
                    iptalturu = 'ABONELİK İPTAL'
                    gerekli = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break

                if  'TARIFEDEGISIMI' in text2:
                    iptalturu = 'TARİFE DEĞİŞİMİ'
                    gerekli = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break
                if 'CIHAZKIRALAMAIPTAL' in text2:
                    iptalturu = 'CİHAZ KİRALAMA İPTAL'
                    gerekli = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break
                if 'CIHAZKIRALAMASIPARIS' in text2:
                    iptalturu = 'CİHAZ KİRALAMA SİPARİŞ'
                    gerekli = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break
                if  'TAAHHUTLUABONELIKDEVIRALMA' in text2:
                    iptalturu = 'TAAHHÜTLÜ ABONELİK DEVİR ALMA'
                    gerekli = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break
                if 'TAAHHUTLUABONELIKDEVIR' in text2:
                    iptalturu = 'TAAHHÜTLÜ ABONELİK DEVİR'
                    gerekli = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break
                if 'TAAHHUTLUYENIABONELIK' in text2:
                    iptalturu = 'TAAHHÜTLÜ YENİ ABONELİK'
                    gerekli = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break
                if  'TAAHHUTLU ABONELIK NAKIL GELEN' and 'TAHHUTLUNAKILGELEN' in text2:
                    iptalturu = 'TAAHHÜTLÜ NAKİL GELEN'
                    gerekli = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break
                if  'HIZMETIASKIYAALMA' in text2:
                    iptalturu = 'HİZMETİ ASKIYA ALMA'
                    gerekli = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break
                if 'CIHAZIADEFORMU' in text2:
                    iptalturu = 'CİHAZ İADE FORMU'
                    gerekli = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break
                if 'KABLONETIPTALBASVURUFORMU' in text2:
                    iptalturu = 'KABLONET İPTAL FORMU'
                    gerekli = False
                    print('')
                    print(
                        '################################################### DEVELOPED BY CAFTELLE ###################################################')
                    print('')
                    print(
                        ' | Müşteri No: ' + musterinofinal + ' | Hizmet No: ' + hizmetnofinal + ' | İş Emri No: ' + isemrinofinal + ' | İş Emri Türü: ' + iptalturu)
                    break

        # Kamera'dan Okunan Yazının ayrıştırılıp Uygun Yere Atanması Kontrolü
        if iptalturu != '(     )':
            savestart = True

        # AYRIŞTIRILMIŞ VERİLERİ EXCELE KAYDETME

        while savestart:

            # Taranan ve Ayrıştırılan Yazının Excel Tablosu Üzerinde Bulunan Sekmelerden Uygun Olana Atanması

            if 'OKUNAMADI' in iptalturu:
                tarih1 = datetime.datetime.now()
                yil1 = tarih1.year
                ay1 = tarih1.month
                gun1 = tarih1.day
                saat1 = tarih1.hour
                dakika1 = tarih1.minute
                toplami1 = str(yil1) + '/' + str(ay1) + '/' + str(gun1) + ' - ' + str(
                    saat1) + ':' + str(dakika1)
                tarihcikti = str(toplami1)

                Ai = 'A' + str(tesisline)
                Bi = 'B' + str(tesisline)
                Ci = 'C' + str(tesisline)
                Di = 'D' + str(tesisline)
                Ei = 'E' + str(tesisline)
                Fi = 'F' + str(tesisline)

                planSheettesis[Ai] = hizmetnofinal
                planSheettesis[Bi] = musterinofinal
                planSheettesis[Ci] = isemrinofinal
                planSheettesis[Di] = 'OKUNAMADI'
                planSheettesis[Ei] = 'ANALOG KABLO TV'
                planSheettesis[Fi] = tarihcikti

                tesisline = tesisline + 1
                savestart = False
                print(' |  B A Ş A R I Y L A   A K T A R I L D I . M A N U E L  D U Z E L T M E  G E R E K I Y O R ')
                beep.beep(5)
                planWorkbook.save('Taranan_Is_Emirleri.xlsx')
                break


            if  'NUMARA TAŞIMALI YENİ ABONELİK' in iptalturu:
                tarih1 = datetime.datetime.now()
                yil1 = tarih1.year
                ay1 = tarih1.month
                gun1 = tarih1.day
                saat1 = tarih1.hour
                dakika1 = tarih1.minute
                toplami1 = str(yil1) + '/' + str(ay1) + '/' + str(gun1) + ' - ' + str(
                    saat1) + ':' + str(dakika1)
                tarihcikti = str(toplami1)

                Ai = 'A' + str(tesisline)
                Bi = 'B' + str(tesisline)
                Ci = 'C' + str(tesisline)
                Di = 'D' + str(tesisline)
                Ei = 'E' + str(tesisline)
                Fi = 'F' + str(tesisline)

                planSheettesis[Ai] = hizmetnofinal
                planSheettesis[Bi] = musterinofinal
                planSheettesis[Ci] = isemrinofinal
                planSheettesis[Di] = iptalturu
                planSheettesis[Ei] = 'KABLO SES'
                planSheettesis[Fi] = tarihcikti

                tesisline = tesisline + 1
                savestart = False
                print(' |  B A Ş A R I Y L A   A K T A R I L D I . ')
                beep.beep(5)
                planWorkbook.save('Taranan_Is_Emirleri.xlsx')
                break

            if  'KABLOSES İPTAL' in iptalturu:
                tarih1 = datetime.datetime.now()
                yil1 = tarih1.year
                ay1 = tarih1.month
                gun1 = tarih1.day
                saat1 = tarih1.hour
                dakika1 = tarih1.minute
                toplami1 = str(yil1) + '/' + str(ay1) + '/' + str(gun1) + ' - ' + str(
                    saat1) + ':' + str(dakika1)
                tarihcikti = str(toplami1)

                Ai = 'A' + str(tesisline)
                Bi = 'B' + str(tesisline)
                Ci = 'C' + str(tesisline)
                Di = 'D' + str(tesisline)
                Ei = 'E' + str(tesisline)
                Fi = 'F' + str(tesisline)

                planSheettesis[Ai] = hizmetnofinal
                planSheettesis[Bi] = musterinofinal
                planSheettesis[Ci] = isemrinofinal
                planSheettesis[Di] = iptalturu
                planSheettesis[Ei] = 'KABLO SES'
                planSheettesis[Fi] = tarihcikti

                tesisline = tesisline + 1
                savestart = False
                print(' |  B A Ş A R I Y L A   A K T A R I L D I . ')
                beep.beep(5)
                planWorkbook.save('Taranan_Is_Emirleri.xlsx')
                break

            if 'TARİFE DEĞİŞİMİ' in iptalturu:
                tarih1 = datetime.datetime.now()
                yil1 = tarih1.year
                ay1 = tarih1.month
                gun1 = tarih1.day
                saat1 = tarih1.hour
                dakika1 = tarih1.minute
                toplami1 = str(yil1) + '/' + str(ay1) + '/' + str(gun1) + ' - ' + str(
                    saat1) + ':' + str(dakika1)
                tarihcikti = str(toplami1)

                Ai = 'A' + str(tesisline)
                Bi = 'B' + str(tesisline)
                Ci = 'C' + str(tesisline)
                Di = 'D' + str(tesisline)
                Ei = 'E' + str(tesisline)
                Fi = 'F' + str(tesisline)

                planSheettesis[Ai] = hizmetnofinal
                planSheettesis[Bi] = musterinofinal
                planSheettesis[Ci] = isemrinofinal
                planSheettesis[Di] = iptalturu
                planSheettesis[Ei] = 'ANALOG KABLO TV'
                planSheettesis[Fi] = tarihcikti

                tesisline = tesisline + 1
                savestart = False
                print(' |  B A Ş A R I Y L A   A K T A R I L D I . ')
                beep.beep(5)
                planWorkbook.save('Taranan_Is_Emirleri.xlsx')
                break

            if 'TAAHHÜTLÜ YENİ ABONELİK' in iptalturu:
                tarih1 = datetime.datetime.now()
                yil1 = tarih1.year
                ay1 = tarih1.month
                gun1 = tarih1.day
                saat1 = tarih1.hour
                dakika1 = tarih1.minute
                toplami1 = str(yil1) + '/' + str(ay1) + '/' + str(gun1) + ' - ' + str(
                    saat1) + ':' + str(dakika1)
                tarihcikti = str(toplami1)

                Ai = 'A' + str(tesisline)
                Bi = 'B' + str(tesisline)
                Ci = 'C' + str(tesisline)
                Di = 'D' + str(tesisline)
                Ei = 'E' + str(tesisline)
                Fi = 'F' + str(tesisline)

                planSheettesis[Ai] = hizmetnofinal
                planSheettesis[Bi] = musterinofinal
                planSheettesis[Ci] = isemrinofinal
                planSheettesis[Di] = iptalturu
                planSheettesis[Ei] = 'ANALOG KABLO TV'
                planSheettesis[Fi] = tarihcikti

                tesisline = tesisline + 1
                savestart = False
                print(' |  B A Ş A R I Y L A   A K T A R I L D I . ')
                beep.beep(5)
                planWorkbook.save('Taranan_Is_Emirleri.xlsx')
                break

            if 'TAAHHÜTLÜ ABONELİK DEVİR' in iptalturu:
                tarih1 = datetime.datetime.now()
                yil1 = tarih1.year
                ay1 = tarih1.month
                gun1 = tarih1.day
                saat1 = tarih1.hour
                dakika1 = tarih1.minute
                toplami1 = str(yil1) + '/' + str(ay1) + '/' + str(gun1) + ' - ' + str(
                    saat1) + ':' + str(dakika1)
                tarihcikti = str(toplami1)

                Ai = 'A' + str(tesisline)
                Bi = 'B' + str(tesisline)
                Ci = 'C' + str(tesisline)
                Di = 'D' + str(tesisline)
                Ei = 'E' + str(tesisline)
                Fi = 'F' + str(tesisline)

                planSheettesis[Ai] = hizmetnofinal
                planSheettesis[Bi] = musterinofinal
                planSheettesis[Ci] = isemrinofinal
                planSheettesis[Di] = iptalturu
                planSheettesis[Ei] = 'ANALOG KABLO TV'
                planSheettesis[Fi] = tarihcikti

                tesisline = tesisline + 1
                savestart = False
                print(' |  B A Ş A R I Y L A   A K T A R I L D I . ')
                beep.beep(5)
                planWorkbook.save('Taranan_Is_Emirleri.xlsx')
                break

            if 'HİZMETİ ASKIYA ALMA' in iptalturu:
                tarih1 = datetime.datetime.now()
                yil1 = tarih1.year
                ay1 = tarih1.month
                gun1 = tarih1.day
                saat1 = tarih1.hour
                dakika1 = tarih1.minute
                toplami1 = str(yil1) + '/' + str(ay1) + '/' + str(gun1) + ' - ' + str(
                    saat1) + ':' + str(dakika1)
                tarihcikti = str(toplami1)

                Ai = 'A' + str(tesisline)
                Bi = 'B' + str(tesisline)
                Ci = 'C' + str(tesisline)
                Di = 'D' + str(tesisline)
                Ei = 'E' + str(tesisline)
                Fi = 'F' + str(tesisline)

                planSheettesis[Ai] = hizmetnofinal
                planSheettesis[Bi] = musterinofinal
                planSheettesis[Ci] = isemrinofinal
                planSheettesis[Di] = iptalturu
                planSheettesis[Ei] = 'ANALOG KABLO TV'
                planSheettesis[Fi] = tarihcikti

                tesisline = tesisline + 1
                savestart = False
                print(' |  B A Ş A R I Y L A   A K T A R I L D I . ')
                beep.beep(5)
                planWorkbook.save('Taranan_Is_Emirleri.xlsx')
                break

            if 'TAAHHÜTLÜ ABONELİK DEVİR ALMA' in iptalturu:
                tarih1 = datetime.datetime.now()
                yil1 = tarih1.year
                ay1 = tarih1.month
                gun1 = tarih1.day
                saat1 = tarih1.hour
                dakika1 = tarih1.minute
                toplami1 = str(yil1) + '/' + str(ay1) + '/' + str(gun1) + ' - ' + str(
                    saat1) + ':' + str(dakika1)
                tarihcikti = str(toplami1)

                Ai = 'A' + str(tesisline)
                Bi = 'B' + str(tesisline)
                Ci = 'C' + str(tesisline)
                Di = 'D' + str(tesisline)
                Ei = 'E' + str(tesisline)
                Fi = 'F' + str(tesisline)

                planSheettesis[Ai] = hizmetnofinal
                planSheettesis[Bi] = musterinofinal
                planSheettesis[Ci] = isemrinofinal
                planSheettesis[Di] = iptalturu
                planSheettesis[Ei] = 'ANALOG KABLO TV'
                planSheettesis[Fi] = tarihcikti

                tesisline = tesisline + 1
                savestart = False
                print(' |  B A Ş A R I Y L A   A K T A R I L D I . ')
                beep.beep(5)
                planWorkbook.save('Taranan_Is_Emirleri.xlsx')

                break

            if 'CİHAZ KİRALAMA SİPARİŞ' in iptalturu:
                tarih1 = datetime.datetime.now()
                yil1 = tarih1.year
                ay1 = tarih1.month
                gun1 = tarih1.day
                saat1 = tarih1.hour
                dakika1 = tarih1.minute
                toplami1 = str(yil1) + '/' + str(ay1) + '/' + str(gun1) + ' - ' + str(
                    saat1) + ':' + str(dakika1)
                tarihcikti = str(toplami1)

                Ai = 'A' + str(tesisline)
                Bi = 'B' + str(tesisline)
                Ci = 'C' + str(tesisline)
                Di = 'D' + str(tesisline)
                Ei = 'E' + str(tesisline)
                Fi = 'F' + str(tesisline)

                planSheettesis[Ai] = hizmetnofinal
                planSheettesis[Bi] = musterinofinal
                planSheettesis[Ci] = isemrinofinal
                planSheettesis[Di] = iptalturu
                planSheettesis[Ei] = 'ANALOG KABLO TV'
                planSheettesis[Fi] = tarihcikti

                tesisline = tesisline + 1
                savestart = False
                print(' |  B A Ş A R I Y L A   A K T A R I L D I . ')
                beep.beep(5)
                planWorkbook.save('Taranan_Is_Emirleri.xlsx')
                break
            if 'CİHAZ KİRALAMA İPTAL' in iptalturu:
                tarih1 = datetime.datetime.now()
                yil1 = tarih1.year
                ay1 = tarih1.month
                gun1 = tarih1.day
                saat1 = tarih1.hour
                dakika1 = tarih1.minute
                toplami1 = str(yil1) + '/' + str(ay1) + '/' + str(gun1) + ' - ' + str(
                    saat1) + ':' + str(dakika1)
                tarihcikti = str(toplami1)

                Ai = 'A' + str(tesisline)
                Bi = 'B' + str(tesisline)
                Ci = 'C' + str(tesisline)
                Di = 'D' + str(tesisline)
                Ei = 'E' + str(tesisline)
                Fi = 'F' + str(tesisline)

                planSheettesis[Ai] = hizmetnofinal
                planSheettesis[Bi] = musterinofinal
                planSheettesis[Ci] = isemrinofinal
                planSheettesis[Di] = iptalturu
                planSheettesis[Ei] = 'ANALOG KABLO TV'
                planSheettesis[Fi] = tarihcikti

                tesisline = tesisline + 1
                savestart = False
                print(' |  B A Ş A R I Y L A   A K T A R I L D I . ')
                beep.beep(5)
                planWorkbook.save('Taranan_Is_Emirleri.xlsx')
                break

            if 'TAAHHÜTLÜ NAKİL GELEN' in iptalturu:
                tarih1 = datetime.datetime.now()
                yil1 = tarih1.year
                ay1 = tarih1.month
                gun1 = tarih1.day
                saat1 = tarih1.hour
                dakika1 = tarih1.minute
                toplami1 = str(yil1) + '/' + str(ay1) + '/' + str(gun1) + ' - ' + str(
                    saat1) + ':' + str(dakika1)
                tarihcikti = str(toplami1)

                Ai = 'A' + str(tesisline)
                Bi = 'B' + str(tesisline)
                Ci = 'C' + str(tesisline)
                Di = 'D' + str(tesisline)
                Ei = 'E' + str(tesisline)
                Fi = 'F' + str(tesisline)

                planSheettesis[Ai] = hizmetnofinal
                planSheettesis[Bi] = musterinofinal
                planSheettesis[Ci] = isemrinofinal
                planSheettesis[Di] = iptalturu
                planSheettesis[Ei] = 'ANALOG KABLO TV'
                planSheettesis[Fi] = tarihcikti

                tesisline = tesisline + 1
                savestart = False
                print(' |  B A Ş A R I Y L A   A K T A R I L D I . ')
                beep.beep(5)
                planWorkbook.save('Taranan_Is_Emirleri.xlsx')
                break

            if 'ABONELİK İPTAL' in iptalturu:
                tarih1 = datetime.datetime.now()
                yil1 = tarih1.year
                ay1 = tarih1.month
                gun1 = tarih1.day
                saat1 = tarih1.hour
                dakika1 = tarih1.minute
                toplami1 = str(yil1) + '/' + str(ay1) + '/' + str(gun1) + ' - ' + str(
                    saat1) + ':' + str(dakika1)
                tarihcikti = str(toplami1)

                Ai = 'A' + str(iptalline)
                Bi = 'B' + str(iptalline)
                Ci = 'C' + str(iptalline)
                Di = 'D' + str(iptalline)
                Ei = 'E' + str(iptalline)
                Fi = 'F' + str(iptalline)

                planSheetiptal[Ai] = hizmetnofinal
                planSheetiptal[Bi] = musterinofinal
                planSheetiptal[Ci] = isemrinofinal
                planSheetiptal[Di] = iptalturu
                planSheetiptal[Ei] = 'ANALOG KABLO TV'
                planSheetiptal[Fi] = tarihcikti

                iptalline = iptalline + 1
                savestart = False
                print(' |  B A Ş A R I Y L A   A K T A R I L D I . ')
                beep.beep(5)
                planWorkbook.save('Taranan_Is_Emirleri.xlsx')
                break

            if 'CİHAZ İADE FORMU' in iptalturu:
                tarih1 = datetime.datetime.now()
                yil1 = tarih1.year
                ay1 = tarih1.month
                gun1 = tarih1.day
                saat1 = tarih1.hour
                dakika1 = tarih1.minute
                toplami1 = str(yil1) + '/' + str(ay1) + '/' + str(gun1) + ' - ' + str(
                    saat1) + ':' + str(dakika1)
                tarihcikti = str(toplami1)

                Ai = 'A' + str(iptalline)
                Bi = 'B' + str(iptalline)
                Ci = 'C' + str(iptalline)
                Di = 'D' + str(iptalline)
                Ei = 'E' + str(iptalline)
                Fi = 'F' + str(iptalline)

                planSheetiptal[Ai] = hizmetnofinal
                planSheetiptal[Bi] = musterinofinal
                planSheetiptal[Ci] = isemrinofinal
                planSheetiptal[Di] = iptalturu
                planSheetiptal[Ei] = 'ANALOG KABLO TV'
                planSheetiptal[Fi] = tarihcikti

                iptalline = iptalline + 1
                savestart = False
                print(' |  B A Ş A R I Y L A   A K T A R I L D I . ')
                beep.beep(5)
                planWorkbook.save('Taranan_Is_Emirleri.xlsx')
                break

            if 'VERASETEN İPTAL' in iptalturu:
                tarih1 = datetime.datetime.now()
                yil1 = tarih1.year
                ay1 = tarih1.month
                gun1 = tarih1.day
                saat1 = tarih1.hour
                dakika1 = tarih1.minute
                toplami1 = str(yil1) + '/' + str(ay1) + '/' + str(gun1) + ' - ' + str(
                    saat1) + ':' + str(dakika1)
                tarihcikti = str(toplami1)

                Ai = 'A' + str(iptalline)
                Bi = 'B' + str(iptalline)
                Ci = 'C' + str(iptalline)
                Di = 'D' + str(iptalline)
                Ei = 'E' + str(iptalline)
                Fi = 'F' + str(iptalline)

                planSheetiptal[Ai] = hizmetnofinal
                planSheetiptal[Bi] = musterinofinal
                planSheetiptal[Ci] = isemrinofinal
                planSheetiptal[Di] = iptalturu
                planSheetiptal[Ei] = 'ANALOG KABLO TV'
                planSheetiptal[Fi] = tarihcikti

                iptalline = iptalline + 1
                savestart = False
                print(' |  B A Ş A R I Y L A   A K T A R I L D I . ')
                beep.beep(5)
                planWorkbook.save('Taranan_Is_Emirleri.xlsx')
                break

            if 'KABLONET İPTAL FORMU' in iptalturu:
                tarih1 = datetime.datetime.now()
                yil1 = tarih1.year
                ay1 = tarih1.month
                gun1 = tarih1.day
                saat1 = tarih1.hour
                dakika1 = tarih1.minute
                toplami1 = str(yil1) + '/' + str(ay1) + '/' + str(gun1) + ' - ' + str(
                    saat1) + ':' + str(dakika1)
                tarihcikti = str(toplami1)

                Ai = 'A' + str(iptalline)
                Bi = 'B' + str(iptalline)
                Ci = 'C' + str(iptalline)
                Di = 'D' + str(iptalline)
                Ei = 'E' + str(iptalline)
                Fi = 'F' + str(iptalline)

                planSheetiptal[Ai] = hizmetnofinal
                planSheetiptal[Bi] = musterinofinal
                planSheetiptal[Ci] = isemrinofinal
                planSheetiptal[Di] = iptalturu
                planSheetiptal[Ei] = 'ANALOG KABLO TV'
                planSheetiptal[Fi] = tarihcikti

                iptalline = iptalline + 1
                tarananadet = tarananadet + 1
                savestart = False
                print(' |  B A Ş A R I Y L A   A K T A R I L D I . ' + 'Taranan Evrak Sayısı: '+ tarananadet)
                beep.beep(5)
                planWorkbook.save('Taranan_Is_Emirleri.xlsx')
                break

        print('')
        print(
            '################################################### DEVELOPED BY CAFTELLE ###################################################')
        print('')
        
        print(" | Belge tarama bittiyse taranan dosyaları Mail ile göndermek için herhangi bir tuşa basıp ENTER tuşuna basınız ve uygulama tekrar başladığında tarama yapmadan Mail gönderme bölümünü seçiniz. ")
       
        a = input(" | Belge Taramaya devam etmek için ENTER`a basınız. ")
        

        if a != "":
            print('')
            print(
                '################################################### DEVELOPED BY CAFTELLE ###################################################')
            print('')
            break




try:    
    
    print(' | [ 1 ] Belge Taramaya başlamak istiyorum. | [ 2 ] Taranan Belgeleri Mail atmak istiyorum. ')
    sec2 = input(' | Lütfen bir seçim yapınız : ')
    sec2 = str(sec2)

    if sec2 == '1':
        AllWithLove()

    elif sec2 == '2':
        MailGonder()
        dosyasilme()
    else:
        print(' | Hatalı kod tuşladınız. Tekrar Deneyiniz.')
        beep.beep(3)
except:
    print(' | Bilinmeyen bir sorun oluştu. Uygulamayı yeniden başlatınız.')
    beep.beep(3)




        
print('')
print('#############################################################################################################################')
print('################################################### DEVELOPED BY CAFTELLE ###################################################')
print('#############################################################################################################################')
print('')

###########################################################################
########################## DEVELOPED BY CAFTELLE ##########################
########################## DEVELOPED BY CAFTELLE ##########################
########################## DEVELOPED BY CAFTELLE ##########################
###########################################################################
################################# I E T ###################################
######################## V E R S I O N  1 . 0 . 5 #########################
###########################################################################
###########################################################################
