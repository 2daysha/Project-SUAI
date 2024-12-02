# Состояния пользователей
USER_STATE = {}

def set_user_state(user_id, state):
    USER_STATE[user_id] = state

def get_user_state(user_id):
    return USER_STATE.get(user_id)
