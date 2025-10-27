import pandas as pd
import re

# Đọc file Excel
df = pd.read_excel("cleaned_data.xlsx")

# Danh sách cặp (từ sai, từ đúng)
replacements = [
    ("báotạm", "báo tạm"),
    ("báotạmvắng", "báo tạm vắng"),
    ("bệnhtạm", "bệnh tạm"),
    ("chấttạmthời", "chất tạm thời"),
    ("côngdân", "công dân"),
    ("cơquan", "cơ quan"),
    ("cưtrú2020", "cư trú 2020"),
    ("cưtrúcó", "cư trú có"),
    ("cưtrúlà", "cư trú là"),
    ("cưtrúthế", "cư trú thế"),
    ("cưtrúthực", "cư trú thực"),
    ("cưtrútrực", "cư trú trực"),
    ("cưtrúở", "cư trú ở"),
    ("gamil", "gmail"),
    ("hạntạm", "hạn tạm"),
    ("hạntạmtrú", "hạn tạm trú"),
    ("họlàm", "họ làm"),
    ("hợpvắngmặt", "hợp vắng mặt"),
    ("khaitạm", "khai tạm"),
    ("khaitạmvắng", "khai tạm vắng"),
    ("kýthường", "ký thường"),
    ("kýtạm", "ký tạm"),
    ("kýtạmtrú", "ký tạm trú"),
    ("kýtạmtrúvà", "ký tạm trú và"),
    ("làmtạm", "làm tạm"),
    ("làmtạmtrú", "làm tạm trú"),
    ("làmtạmtrútạm", "làm tạm trú tạm"),
    ("lưutrú", "lưu trú"),
    ("lưutrúcho", "lưu trú cho"),
    ("lưutrúdu", "lưu trú du"),
    ("lưutrúdài", "lưu trú dài"),
    ("lạitầm2", "lại tầm 2"),
    ("ngoàitạm", "ngoài tạm"),
    ("ngườitạmtrú", "người tạm trú"),
    ("ngườivắngmặt", "người vắng mặt"),
    ("nhậntạm", "nhận tạm"),
    ("nhậntạmtrú", "nhận tạm trú"),
    ("nơitạm", "nơi tạm"),
    ("nơitạmtrú", "nơi tạm trú"),
    ("sổtạm", "sổ tạm"),
    ("thanhtâmđến", "thành tâm đến"),
    ("thườngtrú", "thường trú"),
    ("thườngtrúkhông", "thường trú không"),
    ("thườngtrúluôn", "thường trú luôn"),
    ("thườngtrútrong", "thường trú trong"),
    ("thườngtrútại", "thường trú tại"),
    ("thườngtrúđược", "thường trú được"),
    ("thườngtrúở", "thường trú ở"),
    ("thẻtạm", "thẻ tạm"),
    ("tintạm", "tin tạm"),
    ("tratạmtrú", "tra tạm trú"),
    ("trungtâmthành", "trung tâm thành"),
    ("trú06", "trú 06"),
    ("trúbao", "trú bao"),
    ("trúcho", "trú cho"),
    ("trúcùng", "trú cùng"),
    ("trúcấp", "trú cấp"),
    ("trúcần", "trú cần"),
    ("trúcủa", "trú của"),
    ("trúem", "trú em"),
    ("trúgấp", "trú gấp"),
    ("trúgồm", "trú gồm"),
    ("trúhay", "trú hay"),
    ("trúhiện", "trú hiện"),
    ("trúkhai", "trú khai"),
    ("trúkhi", "trú khi"),
    ("trúkhông", "trú không"),
    ("trúlà", "trú là"),
    ("trúlại", "trú lại"),
    ("trúmới", "trú mới"),
    ("trúnhư", "trú như"),
    ("trúnào", "trú nào"),
    ("trúnơi", "trú nơi"),
    ("trúnếu", "trú nếu"),
    ("trúonline", "trú online"),
    ("trúphải", "trú phải"),
    ("trúqua", "trú qua"),
    ("trúquy", "trú quy"),
    ("trútheo", "trú theo"),
    ("trúthì", "trú thì"),
    ("trúthế", "trú thế"),
    ("trúthực", "trú thực"),
    ("trútrong", "trú trong"),
    ("trútrên", "trú trên"),
    ("trútrực", "trú trực"),
    ("trútại", "trú tại"),
    ("trúvà", "trú và"),
    ("trúvới", "trú với"),
    ("trúđược", "trú được"),
    ("trúđể", "trú để"),
    ("trúđối", "trú đối"),
    ("trúở", "trú ở"),
    ("tạmtrú", "tạm trú"),
    ("vàtạmthời", "và tạm thời"),
    ("vắngbị", "vắng bị"),
    ("vắngcho", "vắng cho"),
    ("vắngcó", "vắng có"),
    ("vắnghay", "vắng hay"),
    ("vắngkhi", "vắng khi"),
    ("vắngkhông", "vắng không"),
    ("vắngnhư", "vắng như"),
    ("vắngonline", "vắng online"),
    ("vắngphải", "vắng phải"),
    ("vắngthì", "vắng thì"),
    ("vắngthông", "vắng thông"),
    ("vắngvà", "vắng và"),
    ("vắngvới", "vắng với"),
    ("vắngđược", "vắng được"),
    ("vắngđối", "vắng đối"),
    ("vắngở", "vắng ở"),
    ("vớitạmtrú", "với tạm trú"),
    ("đangtạm", "đang tạm"),
    ("đótầm3", "đó tầm 3"),
    ("đăngtạm", "đăng tạm"),
    ("được", "được"),
    ("đượctạm", "được tạm"),
    ("địnhtạm", "định tạm"),
    ("nấv", "vấn"),
]

replacements = sorted(replacements, key=lambda x: len(x[0]), reverse=True)

def smart_replace(match, correct_word):
    word = match.group(0)
    if word.isupper():
        return correct_word.upper()
    elif word[0].isupper():
        return correct_word.capitalize()
    else:
        return correct_word

def replace_words(text, replacements):
    if not isinstance(text, str):
        return text
    for wrong, correct in replacements:
        # Thay thế không phân biệt hoa thường
        text = re.sub(
            rf"\b{re.escape(wrong)}\b",
            lambda m: smart_replace(m, correct),
            text,
            flags=re.IGNORECASE
        )
    return text

df['question_note'] = df['question_note'].apply(lambda x: replace_words(x, replacements))
df['Subject'] = df['Subject'].apply(lambda x: replace_words(x, replacements))

df.to_excel("cleaned_data.xlsx", index=False)

print("✅ Đã thay thế xong (có phân biệt chữ hoa/thường) và lưu vào cleaned_data.xlsx")