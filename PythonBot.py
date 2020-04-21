# -*- coding:utf-8 -*- 
import discord
import asyncio # 디스코드 모듈과, 보조 모듈인 asyncio를 불러옵니다.
import datetime
import os
import openpyxl

client = discord.Client() # discord.Client() 같은 긴 단어 대신 client를 사용하겠다는 선언입니다.

@client.event
async def on_ready(): # 봇이 준비가 되면 1회 실행되는 부분입니다.
    # 봇이 "반갑습니다"를 플레이 하게 됩니다.
    # 눈치 채셨을지 모르곘지만, discord.Status.online에서 online을 dnd로 바꾸면 "다른 용무 중", idle로 바꾸면 "자리 비움"으로 바뀝니다.
    await client.change_presence(status=discord.Status.online, activity=discord.Game("반갑습니다 :D"))
    print("I'm Ready!") # I'm Ready! 문구를 출력합니다.
    print(client.user.name) # 봇의 이름을 출력합니다.
    print(client.user.id) # 봇의 Discord 고유 ID를 출력합니다.
@client.event
async def on_message(message): # 메시지가 들어 올 때마다 가동되는 구문입니다.
    if(message.author.bot):
        return
    if  (message.channel.type is discord.ChannelType.private): #and message.author.id!="700608102269059083":#제 봇 아이디
        await message.channel.send('접수완료')
        channel = client.get_channel(558908366739734530)
        await channel.send(message.author.name+" : "+message.content)#관리자에게 메세지가 가는 방식. 아이디는 제꺼
        file = openpyxl.load_workbook("cooltime.xlsx")
        sheet = file.active
        noww=datetime.datetime.now()
        h=noww.hour     
        m=noww.minute
        s=noww.second
        time=int(h)*3600+int(m)*60+int(s)
        for i in range(1, 100):
            if str(sheet["A"+str(i)].value)==str(message.author.id):
                if(int(sheet["B"+str(i)].value)>time-600):
                    await client.send_message(discord.utils.get(client.get_all_members(), id=str(message.author.id)), str(600-time+int(sheet["B"+str(i)].value))+"seconds remaning")
                    return
                else:
                    sheet["B"+str(i)].value=sheet["C"+str(i)].value
                    sheet["C"+str(i)].value=sheet["D"+str(i)].value
                    sheet["D"+str(i)].value=time
                    file.save("cooltime.xlsx")
                    await message.channel.send('접수완료')
                    channel = client.get_channel(558908366739734530)
                    await channel.send(message.author.name+" : "+message.content)#관리자에게 메세지가 가는 방식. 아이디는 제꺼
                    return
        for j in range(1,100):
            if sheet["A"+str(j)].value==None:
                sheet["A"+str(j)].value=message.author.id
                sheet["D"+str(j)].value=str(time)
                channel = client.get_channel(558908366739734530)
                await channel.send(message.author.name+" : "+message.content)#관리자에게 메세지가 가는 방식. 아이디는 제꺼
                file.save("cooltime.xlsx")
                return
access_token=os.environ["BOT_TOKEN"]
client.run(access_token) # 아까 넣어놓은 토큰 가져다가 봇을 실행하라는 부분입니다. 이 코드 없으면 구문이 아무리 완벽해도 실행되지 않습니다.
