

while True:
    prompt_user, prompt_ai, history = '', '', ''
    while True:
        prompt_user = input('用户:')
        if prompt_user == 'done':
            with open('conversations_1.txt', 'at', encoding='utf-8') as f:
                spl = history.split('\\nChatSW:')
                f.write('"{}\\nChatSW:","{}",[],generate\n'.format('\\nChatSW:'.join(spl[:-1]), spl[-1]))
            history = ''
            print('[对话已重置]')
            break
        if prompt_user and prompt_user != 'done':
            history += '用户:{}\\n'.format(prompt_user)
        prompt_ai = input('ChatSW:')
        if prompt_ai and prompt_user != 'done':
            history += 'ChatSW:{}\\n'.format(prompt_ai)
