import pypresence
import os
import win32com.client
import time



if __name__=='__main__':
    with open("Secret.txt","r") as f:
        secret = f.read()

    o = win32com.client.Dispatch("iTunes.Application") #connect to the COM of iTunes.Application
    DiscordRPC = pypresence.Presence(secret, pipe=0) #setup rich presence
    DiscordRPC.connect() #connect to it
    #Get all relevant information
    track = o.CurrentTrack.Name
    artist = o.CurrentTrack.Artist
    state = o.PlayerState
    #timestamps for computing how far into the song we are
    starttime = int(time.time()) - o.PlayerPosition
    endtime = int(time.time()) + (o.CurrentTrack.Duration - o.PlayerPosition)
    #ship it off
    DiscordRPC.update(state=artist, details=track, start=starttime, end=endtime)
    while True:
        #Has our track changed?
        if track == o.CurrentTrack.Name:
            time.sleep(5) #sleep to check again in 5 seconds
            continue
        #Since the track isn't what's currently playing we want to gather information again
        track = o.CurrentTrack.Name
        artist = o.CurrentTrack.Artist
        state = o.PlayerState
        #timestamps for computing how far into the song we are
        starttime =   int(time.time()) - o.PlayerPosition
        endtime = int(time.time()) + (o.CurrentTrack.Duration - o.PlayerPosition)
        #ship it
        DiscordRPC.update(state=artist,details=track,start=starttime,end=endtime)
        #Sleep to check again to avoid discord rate limits and be nice to their servers
        time.sleep(5)






