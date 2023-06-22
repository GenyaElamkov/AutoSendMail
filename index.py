"""
Скрипт в автоматическом режиме, 1 раз в месяц отправляет PDF файл на почту.
0 10 6 * ? *
"""

from main import main


def start(event, context):
    main()

    return {
        'statusCode': 200,
        'body': 'Сообщение отправлено',
    }
