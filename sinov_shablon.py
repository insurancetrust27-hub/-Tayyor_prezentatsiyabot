"""Sinov: AI'siz, tayyor namuna matn bilan shablonni to'ldiradi va sinov.pptx yaratadi."""
from shablon import shablondan_yaratish

NAMUNA = {
    "muqova_1": "Quyosh", "muqova_2": "tizimi", "muqova_izoh": "Yulduzimiz va uning sayyoralari",
    "s2_a": "Quyosh tizimi Quyosh va uning atrofida aylanuvchi sayyoralar, yo'ldosh va kichik jismlardan iborat. U taxminan 4,6 milliard yil oldin gaz va chang bulutidan paydo bo'lgan.",
    "s2_b": "Tizim markazida Quyosh joylashgan va u butun tizim massasining deyarli hammasini tashkil qiladi. Sayyoralar Quyosh atrofida ellips shaklidagi orbitalar bo'ylab harakatlanadi.",
    "s3_a": "Quyosh tizimini o'rganish Yerning kelib chiqishini va koinotdagi o'rnimizni tushunishga yordam beradi.",
    "s3_b": "Sayyoralarni kuzatish orqali olimlar iqlim, atmosfera va tortishish kuchi qonuniyatlarini yaxshiroq bilib oladi. Bu bilimlar kosmik parvozlar va sun'iy yo'ldoshlar uchun ham muhim.",
    "s4_a": "Qadimgi davrlardan beri odamlar osmonni kuzatib, sayyoralarning harakatini yozib borgan.",
    "s4_b": "Bugun tizimda sakkizta sayyora bor: Merkuriy, Venera, Yer, Mars, Yupiter, Saturn, Uran va Neptun. Ular ikki guruhga bo'linadi: ichki tosh sayyoralar va tashqi gaz gigantlari.",
    "s5_a": "Kosmik zondlar va teleskoplar tizimni tobora chuqurroq o'rganishga imkon beryapti.",
    "s5_b": "Zondlar sayyoralar yonidan o'tib, ularning suratlari va ma'lumotlarini Yerga yuboradi. Marsda o'rganish uchun maxsus roverlar ishlaydi. Bu tadqiqotlar hayot izlarini qidirishga yordam beradi.",
    "s6": "Sayyoralar o'lchami, tarkibi va Quyoshdan uzoqligi bilan farq qiladi. Yerga yaqin sayyoralar toshdan, uzoqdagilar asosan gazdan iborat.",
    "s7_a": "Kosmosni o'rganish katta xarajat va murakkab texnologiyani talab qiladi.",
    "s7_b": "Masofalar juda katta, shuning uchun zondlar sayyoralarga yillar davomida uchadi. Radiatsiya va sovuq uskunalarga zarar yetkazadi. Shunga qaramay, xalqaro hamkorlik bu muammolarni kamaytiryapti.",
    "s8": "Quyosh tizimi haqidagi bilimlar texnologiyani rivojlantiradi, yangi kashfiyotlarga yo'l ochadi va odamlarni ilm-fanga qiziqtiradi.",
    "s9_a": "Orbita - jismning boshqa jism atrofidagi harakat yo'li. Sayyora - Quyosh atrofida aylanuvchi katta jism. Yo'ldosh - sayyora atrofida aylanuvchi jism. Asteroid - toshli kichik jism.",
    "s9_b": "Kometa - muz va changdan iborat jism. Galaktika - yulduzlar to'plami. Yorug'lik yili - yorug'lik bir yilda bosib o'tadigan masofa.",
    "s10_t1": "Sayyora", "s10_t2": "turlari",
    "s10_k1": "Tosh sayyoralar", "s10_k1_m": "Merkuriy, Venera, Yer va Mars qattiq yuzaga ega.",
    "s10_k2": "Gaz gigantlar", "s10_k2_m": "Yupiter va Saturn asosan vodorod va geliydan iborat.",
    "s10_k3": "Muz gigantlar", "s10_k3_m": "Uran va Neptun tarkibida muzlagan moddalar ko'p.",
    "s10_k4": "Mitti sayyoralar", "s10_k4_m": "Pluton kabi kichik jismlar mitti sayyora hisoblanadi.",
    "s11_t1": "Kichik", "s11_t2": "jismlar",
    "s11_k1": "Asteroidlar", "s11_k1_m": "Mars va Yupiter orasidagi kamarda joylashgan toshli jismlar.",
    "s11_k2": "Kometalar", "s11_k2_m": "Quyoshga yaqinlashganda dumli shaklga kiradigan muzli jismlar.",
    "s11_k3": "Meteoritlar", "s11_k3_m": "Yer atmosferasiga kirib, yerga tushgan kosmik jismlar.",
    "s12_t1": "Quyosh", "s12_t2": "energiyasi", "s12_sub": "Hayot manbai",
    "s12_a": "Quyosh yorug'lik va issiqlik chiqaradi, bu Yerdagi hayot uchun asosiy manba.",
    "s12_b": "Quyosh energiyasidan elektr olish uchun quyosh panellaridan foydalaniladi.",
    "s13_t1": "Yer", "s13_t2": "sayyorasi",
    "s13_a": "Yer Quyoshdan uchinchi sayyora va hayot ma'lum bo'lgan yagona joy.",
    "s13_b": "Atmosfera va suv Yerni hayot uchun qulay qiladi.",
    "s14_t1": "Kelajak", "s14_t2": "kosmosi",
    "s14_k1": "Marsga parvoz", "s14_k1_m": "Kelajakda odamlarni Marsga yuborish rejalashtirilmoqda.",
    "s14_k2": "Kosmik turizm", "s14_k2_m": "Xususiy kompaniyalar oddiy odamlar uchun kosmik sayohat tashkil qilmoqda.",
    "yakun_izoh": "Koinot sirlari hali ochilmagan",
}

if __name__ == "__main__":
    fayl = shablondan_yaratish("moda", NAMUNA)
    with open("sinov.pptx", "wb") as f:
        f.write(fayl.read())
    print("sinov.pptx tayyor")
