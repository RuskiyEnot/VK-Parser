# VK-Parser
Парсер комментариев к постам сообщества Вконтакте
Собирает все комментарии к N-ому числу постов в каком-либо сообществе ВК.
Результаты выводит в таблицу Excel в виде ID | Имя | Фамилия | Комментарий.

ГАЙД
Вводите в командную строку
1. <python COMMENT_VK.py>

2. Enter the VK group domain:
тут вводите имя сообщества после "https://vk.com/.."
Например: apiclub

3. Enter your VK API access token:
Вводите свой acess token.

4. Enter the number of posts to process:
Вводите количество постов
Например: 100

Результат будет сохранён в файл экселя (если у вас больше 50 постов - то файлы будут сохраняться в несколько этапов по 50, а в конце объединяться в один общий)

P.S. Я не программист делал для своих целей, моим целям вполне подошёл, Вы можете его улучшить или использовать таким какой он есть
