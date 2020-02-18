import outlook
import config
mail = outlook.Outlook()
mail.login(config.UserName,config.Password)
mail.inbox()
print(mail.unread())


