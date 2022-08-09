class NickName:
    def nickname(self, companies):
        nicknames = {}

        for company in companies:

            words_in_name = company.split(' ')

            nickname = ''

            if len(words_in_name[0]) >= 4 and len(words_in_name[0]) <= 8: # if the first word of the company is at least 4 but no greater than 8 characters, make that the nickname
                nickname = words_in_name[0]
            elif len(words_in_name) >= 4: # if there are at least 5 words in the name, use the first letter of each word as the nickname
                for word in words_in_name:
                    nickname += word[0]
                    if len(nickname) == 8:
                        break
            else: # otherwise, use the first 8 characters of the first word of the name
                nickname = words_in_name[0][0:8]

            nickname = nickname.upper()
            nickname = ''.join(ch for ch in nickname if ch.isalnum())

            if nickname in nicknames:
                # while True:
                #     nickname = input(f"nickname {nickname} for company '{company}' conflicts with nickname for company '{nicknames[nickname]}', please enter another nickname: ")
                #     if len(nickname) > 8:
                #         print("nickname too long! please enter a nickname (IPU code) that's 8 characters or less")
                #     else:
                #         break
                nickname = nickname[0:7] + words_in_name[1][0]
                
                nickname = nickname.upper()
                nicknames[nickname] = company
            else:
                nicknames[nickname] = company

        return nicknames

if __name__ == "__main__":
    nicknamer = NickName()
    companies = ["SoftBank Corp.", "SoftBank Corporation"]
    print(nicknamer.nickname(companies))