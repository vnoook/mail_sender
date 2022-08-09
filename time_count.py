letters_all = 16
letters_pack = 3
delay_letter = 1
delay_pack = 2

def time_count(letters_all=16, letters_pack=5, delay_letter=3, delay_pack=300):
    q_full_pack = letters_all // letters_pack

    if letters_all % letters_pack == 0:
        time_pack = (q_full_pack - 1) * delay_pack
        time_letters = (((letters_pack - 1) * delay_letter) * q_full_pack)
        time_short_pack = 0
    else:
        time_pack = q_full_pack * delay_pack
        time_letters = (((letters_pack - 1) * delay_letter) * q_full_pack)
        time_short_pack = ((letters_all % letters_pack) - 1) * delay_letter

    time_all = time_pack + time_letters + time_short_pack

    return time_all

print(f'{time_count(letters_all, letters_pack, delay_letter, delay_pack)}')
