from docxtpl import DocxTemplate, InlineImage
from faker import Faker

fake = Faker('ru_RU')

template = DocxTemplate('../lab2/lab2_template.docx')

n = 1

sing_img_path = 'signature.png'
stamp_img_path = 'stamp.png'

sing_img = InlineImage(template, sing_img_path)
stamp_img = InlineImage(template, stamp_img_path)

for i in range(n):
    context = {
        'initials': fake.name_male(),
        'address': fake.address(),
        'date': fake.date(),
        'time': fake.time(),
        'military_address': fake.address(),
        'signature': sing_img,
        'stamp': stamp_img,
    }
    template.render(context)
    template.save(f"{context['initials'].replace(' ', '_')}.docx")