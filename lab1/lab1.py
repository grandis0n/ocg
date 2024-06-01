from faker import Faker
import random
from docxtpl import DocxTemplate

fake = Faker('ru_RU')


def products_generator():
    products = []
    for _ in range(15):
        amount = random.randint(1, 1000)
        price = random.randint(15, 100000)
        product = {
            'title': fake.word(),
            'code': str(random.randint(1000000, 10000000)),
            'unit': random.choice(['шт', 'л', 'кг']),
            'amount': amount,
            'price': price,
            'sum': price * amount
        }
        products.append(product)
    return products


def generator():
    products = products_generator()
    context = {
        'company': fake.company(),
        'check_number': str(random.randint(1, 10000000)),
        'day': str(random.randint(1, 30)),
        'month': fake.month_name(),
        'year': str(random.randint(10, 24)),
        'seller': fake.name(),
        'address': fake.address(),
        'ORGN': fake.businesses_ogrn(),
        'products': products,
        'general_sum': str(sum(product['sum'] for product in products)) + ' рублей'
    }
    return context


checks = 1

contexts = [generator() for _ in range(checks)]

template = DocxTemplate("tmp.docx")

for i in range(checks):
    template.render(contexts[i])
    template.save(f'././check{i + 1}.docx')
