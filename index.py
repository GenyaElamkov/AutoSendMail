"""
Скрипт в автоматическом режиме 1 раз в месяц отправляет PDF файл на почту.
"""

from main import main


def start(event, context):
    main()

    return {
        'statusCode': 200,
        'body': 'Сообщение отправлено',
    }
