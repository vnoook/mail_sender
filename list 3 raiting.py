q_all_letters = 17
q_letter = 5
mail_delay = 3
mailpack_delay = 300

list1 = [x for x in range(1, q_all_letters+1)]
print(list1)

for i in range(0, len(list1), q_letter):
    list2 = list1[i:i+q_letter]
    print(list2, end=' ')
print()
print('*'*70)

for i in range(0, len(list1), q_letter):
    list2 = list1[i:i+q_letter]
    print(list2)
    for j in list2:
        print(j)
        if list2.index(j) != len(list2)-1:
            print('задержка в секундах', mail_delay)

    if len(list2) == q_letter:
        if q_all_letters not in list2:
            print()
            print('задержка в секундах', mailpack_delay)
    print()
