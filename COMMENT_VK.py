import requests
import re
import demoji
import time
import json
import logging
from openpyxl import Workbook
from tqdm import tqdm

# Функция для получения ID группы по её доменному имени
def get_group_id(group_domain, access_token):
    base_url = 'https://api.vk.com/method/groups.getById'
    params = {
        'group_id': group_domain,
        'access_token': access_token,
        'v': '5.131'
    }

    try:
        response = requests.get(base_url, params=params)
        response.raise_for_status()
        data = response.json()
        group_id = data['response'][0]['id']
        return group_id
    except requests.RequestException as e:
        logging.error(f"Error occurred while fetching group ID: {e}")
        return None
    except (KeyError, IndexError) as e:
        logging.error(f"Error parsing group ID from response: {e}")
        return None

# Функция для получения постов из группы VK
def get_posts_with_comments(group_id, access_token, count=100):
    base_url = 'https://api.vk.com/method/wall.get'
    posts_with_comments = []

    # Параметры для первого запроса
    params = {
        'owner_id': -group_id,
        'count': min(count, 100),
        'offset': 0,
        'access_token': access_token,
        'v': '5.131'
    }

    try:
        response = requests.get(base_url, params=params)
        response.raise_for_status()
        data = response.json()

        if 'response' in data and 'items' in data['response']:
            items = data['response']['items']
            for post in items:
                if post.get('comments', {}).get('count', 0) > 0:  # Проверяем, есть ли у поста комментарии
                    posts_with_comments.append(post)
                    if len(posts_with_comments) == count:
                        break

            # Если запрошенных постов меньше, чем count, то дальше делать запросы не нужно
            count -= len(posts_with_comments)

            # Если count все еще больше нуля, делаем дополнительные запросы с учетом offset
            offset = len(items)
            while count > 0:
                params['offset'] = offset
                params['count'] = min(count, 100)

                response = requests.get(base_url, params=params)
                response.raise_for_status()
                data = response.json()

                if 'response' in data and 'items' in data['response']:
                    items = data['response']['items']
                    for post in items:
                        if post.get('comments', {}).get('count', 0) > 0:
                            posts_with_comments.append(post)
                            count -= 1
                            if count == 0:
                                break
                    offset += len(items)
                else:
                    logging.warning('No posts found in the response.')
                    break

                # Выводим информацию о прогрессе
                logging.info(f"Processed {len(posts_with_comments)} posts with comments. {count} remaining.")

                time.sleep(1)  # Добавляем небольшую задержку между запросами
        else:
            logging.warning('No posts found in the response.')
    except requests.RequestException as e:
        logging.error(f"Error occurred while fetching posts: {e}")
    except json.JSONDecodeError as e:
        logging.error(f"Error decoding JSON: {e}")

    return posts_with_comments

# Остальной код здесь

# Функция для сбора пользователей, оставивших комментарии
def get_comment_users(group_id, access_token, post_ids):
    base_url = 'https://api.vk.com/method/wall.getComments'
    users = []

    for post_id in post_ids:
        params = {
            'owner_id': -group_id,
            'post_id': post_id,
            'count': 100,  # Максимальное количество комментариев за один запрос
            'access_token': access_token,
            'v': '5.131'
        }

        try:
            response = requests.get(base_url, params=params)
            response.raise_for_status()
            data = response.json()

            if 'response' in data and 'items' in data['response']:
                items = data['response']['items']
                for item in items:
                    user_id = item['from_id']
                    comment_text = item.get('text', '')
                    first_name, last_name = get_user_info(user_id, access_token)  # Получаем имя и фамилию пользователя
                    if first_name is not None and last_name is not None:
                        users.append((user_id, first_name, last_name, comment_text))
                    else:
                        logging.warning(f"No user info found for user ID: {user_id}")
            else:
                logging.warning(f'No comments found for post ID: {post_id}')
        except requests.RequestException as e:
            logging.error(f"Error occurred while fetching comments: {e}")
        except json.JSONDecodeError as e:
            logging.error(f"Error decoding JSON: {e}")

        time.sleep(0.35)  # Добавляем небольшую задержку между запросами

    return users
# Функция для получения информации о пользователе (имя и фамилия)
def get_user_info(user_id, access_token):
    base_url = 'https://api.vk.com/method/users.get'
    params = {
        'user_ids': user_id,
        'access_token': access_token,
        'v': '5.131',
        'fields': 'first_name,last_name'  # Указываем, что нам нужны имя и фамилия пользователя
    }

    try:
        response = requests.get(base_url, params=params)
        response.raise_for_status()
        data = response.json()
        if 'response' in data and data['response']:
            user_info = data['response'][0]  # Получаем информацию о пользователе
            return user_info.get('first_name', ''), user_info.get('last_name', '')
        else:
            logging.warning(f"No user info found for user ID: {user_id}")
            return None, None
    except requests.RequestException as e:
        logging.error(f"Error occurred while fetching user info: {e}")
        return None, None

# Функция для сохранения пользовательских ID, имени, фамилии и комментариев в файл Excel
def save_comment_users_to_excel(users, filename):
    wb = Workbook()
    ws = wb.active
    ws.append(['ID Пользователя', 'Имя', 'Фамилия', 'Комментарий'])

    for user_id, first_name, last_name, comment_text in users:
        ws.append([user_id, first_name, last_name, comment_text])

    wb.save(filename)

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    
    # Введите домен группы и токен доступа VK API
    group_domain = input('Enter the VK group domain: ')
    access_token = input('Enter your VK API access token: ')

    # Введите количество постов для обработки
    count = int(input('Enter the number of posts to process: '))

    filename = f'{group_domain}_comment_users.xlsx'

    logging.info('Please wait...\n')

    group_id = get_group_id(group_domain, access_token)

if group_id is not None:
    posts_with_comments = get_posts_with_comments(group_id, access_token, count)

    if posts_with_comments:
        # Разделение постов на пакеты по 50
        post_batches = [posts_with_comments[i:i+50] for i in range(0, len(posts_with_comments), 50)]

        # Инициализация общего списка пользователей и комментариев
        all_comment_users = []

        # Обработка каждого пакета постов
        for batch_num, post_batch in enumerate(post_batches, start=1):
            logging.info(f"Processing batch {batch_num}/{len(post_batches)}")

            # Получение пользователей и комментариев для текущего пакета
            comment_users_batch = []
            for post in post_batch:
                post_id = post['id']
                comment_users = get_comment_users(group_id, access_token, [post_id])
                comment_users_batch.extend(comment_users)

            # Добавление пользователей и комментариев текущего пакета к общему списку
            all_comment_users.extend(comment_users_batch)

            # Сохранение результатов текущего пакета в файл
            save_comment_users_to_excel(comment_users_batch, f"{filename}_{batch_num}.xlsx")

        # Сохранение всех результатов в один файл
        save_comment_users_to_excel(all_comment_users, filename)

        logging.info(f'User IDs, names, and comments who commented on posts in group {group_domain} have been saved to {filename}.')
    else:
        logging.warning('No posts with comments fetched for processing.')
else:
    logging.error('Failed to fetch group ID. Check if the group domain and access token are correct.')

