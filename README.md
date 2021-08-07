# How-to-create-a-voice-assistantin-notepad-Searching-google-and-amazon-through-it-source-code-
Set Sapi = Wscript.CreateObject("SAPI.SpVoice")
set wshshell = wscript.CreateObject("wscript.shell")

dim Input

wshshell.run "%windir%\Speech\Common\sapisvr.exe -SpeechUX"
Sapi.speak "Welcome to your pc master please search what do you want to open?"
Input = InputBox("Please type your command here")

if Input = "hello" OR Input = "HELLO" then
Sapi.Speak "Hello my master!I am LEXI your voice assistant"
wshshell.run "LEXI.vbs"

else
if Input = "Hi lexi what can you do for me?" then
Sapi.Speak "Hello My name is LEXI your Virtual assistant My work is to serve you and I can open any applications in your computer"
wshshell.run "LEXI.vbs"

else
If Input =  "calculator" OR Input = "Lexi open calcuator" then
Sapi.speak "Opening calculator"
wshshell.run "calc"
wshshell.run "LEXI.vbs"

else
If Input = "youtube" OR Input = "Lexi open youtube" then
Sapi.speak "Opening youtube"
wshshell.run "www.youtube.com"
wshshell.run "LEXI.vbs"

else
If Input = "mail" OR Input = "Lexi open gmail" then
Sapi.speak "Opening gmail"
wshshell.run "https://mail.google.com"
wshshell.run "LEXI.vbs"

else
If Input = "cp" OR Input = "open cp" OR Input = "Lexi open control panel" then
Sapi.speak "Opening control panel"
wshshell.run "Control Panel"
wshshell.run "LEXI.vbs"

else
if Input = "google" OR Input = "Google" OR Input = "Lexi open google" then
Sapi.speak "Opening google"
wshshell.run "www.google.com"
wshshell.run "LEXI.vbs"

else
if Input = "notepad" OR Input = "Lexi open notepad" then
Sapi.speak "Opening notepad"
wshshell.run "Notepad"
wshshell.run "LEXI.vbs"

else
if Input = "edge" OR Input = "Lexi open edge" then
Sapi.speak "Opening microsoft edge"
wshshell.run "https://www.bing.com/search?q=microsoft+edge&form=EDGHPC&qs=PF&cvid=22c1b82e96d943968aaebc608917148d&cc=US&setlang=en-US"
wshshell.run "LEXI.vbs"

else
if Input = "chrome" OR Input = "Lexi open chrome" then
Sapi.speak "Opening chrome"
wshshell.run "chrome"
wshshell.run "LEXI.vbs"

else
if Input = "internet explorer" OR Input = "Lexi open internet explorer" then
Sapi.speak "Opening Internet Explorer"
wshshell.run "iexplore"
wshshell.run "LEXI.vbs"

else
if Input = "wordpad" OR Input = "Lexi open wordpad" then
Sapi.speak "Opening WordPad"
wshshell.run "WordPad"
wshshell.run "LEXI.vbs"

else
if Input = "netflix" OR Input = "Lexi open netflix" then
Sapi.speak "Opening Netflix"
wshshell.run "https://www.netflix.com/browse"
wshshell.run "LEXI.vbs"

else
if Input = "cmp" OR Input = "Lexi open command prompt" then
Sapi.speak "Opening comand prompt"
wshshell.run "cmd"
wshshell.run "LEXI.vbs"

else
if Input = "Lexi give me some food suggestions" OR Input = "give me some food suggestions" then
Sapi.speak "Here! I have some food suggestions for you like Chinese, Indian, Italian, Tibetan, Japanese or some food items like Chips or nooddles,In some Italian you can eat Pizza and Pasta etc.Thank you"
wshshell.run "LEXI.vbs"

else
if Input = "Lexi play some hindi songs" OR Input  = "play some hindi songs" then
Sapi.speak "Sure playing a list of hindi songs on youtube music"
wshshell.run "https://music.youtube.com/watch?v=fdubeMFwuGs&list=PLKec4N0CFeBnjYLZy_Drg3O9hFWJ8KuTB"
wshshell.run "LEXI.vbs"

else
if Input = "Lexi play some english songs" OR Input  = "play some songs" then
Sapi.speak "Sure playing a list of english songs on youtube music"
wshshell.run "https://music.youtube.com/watch?v=sA_1LZ-cM6c&list=PLKec4N0CFeBnTbLuNPABTn5hmUVP2mU6F"
wshshell.run "LEXI.vbs"

else
if Input = "Lexi can you suggest me some names of hindi movies?" OR Input = "suggest me some hindi movies" then
Sapi.speak "Sure I have list of hindi movies from google we have Mumbai saga, 3 IDIOTS, Coolie no.1, LunchBox, Hotel mumbai, Dhoom 3 and more"
wshshell.run "https://www.bing.com/search?q=list%20me%20some%20hindi%20movies&toWww=1&redig=593A6089676B450694F663F595EAB16A"
wshshell.run "LEXI.vbs"

else
if Input = "Lexi,I want to read news articles" OR Input  = "open a news website for reading articles" then
Sapi.speak "Sure here I have a website for news articles"
wshshell.run "https://timesofindia.indiatimes.com/defaultinterstitial.cms"
wshshell.run "LEXI.vbs"

else
if Input = "hotstar" OR Input = "Lexi open hotstar" then
Sapi.speak "Opening hotstar"
wshshell.run "https://www.hotstar.com/in"
wshshell.run "LEXI.vbs"

else
if Input = "prime video" OR Input = "Lexi open prime video" then
Sapi.speak "Opening Amazon prime video"
wshshell.run "https://www.primevideo.com/?ref_=dvm_pds_amz_in_as_s_b_brand27|m_4mEHKPKZc_c"
wshshell.run "LEXI.vbs"

else
if Input = "Lexi suggest me some english movies" OR Input = "suggest me some movies" then
Sapi.speak "Sure! I have a list of movies from google we have movies like Tenet, enola holmes from netflix, The devil all the time, Black panther, fifty shades of grey, The old guard from netflix, the new mutants, knives out and more.Thank you"
wshshell.run "https://www.bing.com/search?q=list%20me%20some%20hollywood%20movies&toWww=1&redig=ACB5F31BF1054481888CB3F403FA7F8C"
wshshell.run "LEXI.vbs"

else
if Input = "files" OR Input = "Lexi open file explorer" OR Input = "open file explorer" OR Input = "file explorer" then
Sapi.speak "opening file explorer"
wshshell.run "explorer"
wshshell.run "LEXI.vbs"

else
 if Input = "" then
 Sapi.speak "No match found"
 wshshell.run "LEXI.vbs"

else
If Input = "zoom" OR Input = " Lexi open zoom" OR Input = "open zoom" then
Sapi.speak "opening zoom on google"
wshshell.run "https://zoom.us/"
wshshell.run "LEXI.vbs"

else
If Input = "classroom" OR Input = " Lexi open classroom" OR Input = "open classroom" then
Sapi.speak "opening google classroom on google"
wshshell.run "https://classroom.google.com/u/0/h"
wshshell.run "LEXI.vbs"

else
If Input = "meet" OR Input = " Lexi open meet" OR Input = "open meet" then
Sapi.speak "opening google meet on google"
wshshell.run "https://apps.google.com/meet/"
wshshell.run "LEXI.vbs"

else
If Input = "time" OR Input = " Lexi what is the time" OR Input = "time today" then
Sapi.speak "The Time is " & hour(time) & O'clock"
wshshell.run "LEXI.vbs"

else
If Input = "excel" OR Input = "Lexi open excel" OR Input = "open excel" then
Sapi.speak "Opening Microsoft Office Excel 2007"
wshshell.run "Excel"
wshshell.run "LEXI.vbs"

else 
if Input = "files" OR Input = "Lexi open files" OR Input = "open files" then
Sapi.speak "Opening Files Explorer"
wshshell.run "explorer"
wshshell.run "LEXI.VBS"

else
if Input = "recorder" OR Input = "Lexi open sound recorder" then
Sapi.speak "Opening Sound Recorder"
wshshell.run "SoundRecorder"
wshshell.run "LEXI.VBS"

else
if Input = "sync center" OR Input = "Lexi open sync center" then
Sapi.speak "Opening sync center"
wshshell.run "mobsync"
wshshell.run "LEXI.VBS"

else
if Input = "windows update" OR Input = "Lexi update my system" then
Sapi.speak "Opening windows update"
wshshell.run "wuapp"
wshshell.run "LEXI.VBS"

else
if Input = " music player" OR Input = "Lexi open windows media player" then
Sapi.speak "Opening windows media player"
wshshell.run "wmplayer"
wshshell.run "LEXI.VBS"

else
If Input = "On screen keyboard" OR Input = "Lexi open on screen keyboard" then
Sapi.speak "Open on-screen-keyboard"
wshshell.run "osk"
wshshell.run "LEXI.VBS"

else
If Input = "shut down" OR Input = "Lexi shut down the computer" then
Sapi.speak "Shutting down your computer"
wshshell.run "shutdown /s /t 0"
wshshell.run "LEXI.vbs"

else
if Input = "paint" OR Input = "Lexi open paint" then
Sapi.speak "opening paint"
wshshell.run "Paint"
wshshell.run "LEXI.VBS"

else
if Input = "instagram" OR Input = "Lexi open instagram" OR Input = "INSTAGRAM" then
Sapi.speak "opening instagram"
wshshell.run "https://www.instagram.com/"
wshshell.run "LEXI.VBS"

else
if Input = "amazon" OR Input = "Lexi open amazon" then
Sapi.speak "opening amazon"
wshshell.run "https://www.amazon.in/"
wshshell.run "LEXI.VBS"

else
if Input = "discord" OR Input = "Lexi open discord" then
Sapi.speak "opening discord"
wshshell.run "https://discord.com/"
wshshell.run "LEXI.VBS"

else
if Input = "flipkart" OR Input = "Lexi open flipkart" then
Sapi.speak "opening flipkart"
wshshell.run "https://www.flipkart.com/"
wshshell.run "LEXI.VBS"

else
if Input = "facebook" OR Input = "Lexi open facebook" then
Sapi.speak "opening facebook"
wshshell.run "https://www.facebook.com/"
wshshell.run "LEXI.VBS"

else
if Input = "snapchat" OR Input = "Lexi open snapchat" then
Sapi.speak "opening snapchat"
wshshell.run "https://www.snapchat.com/"
wshshell.run "LEXI.VBS"

else
if Input = "settings" OR Input = "Lexi open settings" then
Sapi.speak "opening settings"
wshshell.run "settings /s -t 0"
wshshell.run "LEXI.VBS"

else
if Input = "store" OR Input = "Lexi open store" then
Sapi.speak "opening store"
wshshell.run "store"
wshshell.run "LEXI.VBS"

else
if Input = "restart" OR Input = "Lexi restart my computer" then
Sapi.speak "opening restarting your computer"
wshshell.run "shutdown /r -t 0"
wshshell.run "LEXI.VBS"

else
if Input = "log out" OR Input = "Lexi log out of my computer" then
Sapi.speak "logging you out of your computer"
wshshell.run "shutdown /l"
wshshell.run "LEXI.VBS"

else
if Input = "sleep" OR Input = "Lexi sleep down my computer" then
Sapi.speak "sleeping down your computer"
wshshell.run "rundll32.exe powrprof.dlll. SetSuspendState 0.1.0 "
wshshell.run "LEXI.VBS"

else
if Input = "open a website for downloading latest movies" OR Input = "Lexi open a website for downloading latest movies" then
Sapi.speak "Sure here I have a website for downloading english movies from hollywood"
wshshell.run "http://103.222.20.150/ftpdata/Movies/Hollywood/2020/"
wshshell.run "LEXI.VBS"

else
if Input = "open a website for downloading latest hindi movies" OR Input = "Lexi open a website for downloading hindi latest movies" then
Sapi.speak "Sure here I have a website for downloading Hindi movies from bollywood"
wshshell.run "http://103.222.20.150/ftpdata/Movies/Bollywood/2021/"
wshshell.run "LEXI.VBS"

else
if Input = "open a website for downloading latest Tamil movies" OR Input = "Lexi open a website for downloading tamil latest movies" then
Sapi.speak "Sure here I have a website for downloading Tamil movies from Tollywood"
wshshell.run "http://103.222.20.150/ftpdata/Movies/Tamil/South%20Indian%20Movies%20%282020%29/"
wshshell.run "LEXI.VBS"

else
if Input = "open a website for downloading latest Indian bangla movies" OR Input = "Lexi open a website for downloading Indian bangla latest movies" then
Sapi.speak "Sure here I have a website for downloading Indian bangla movies"
wshshell.run "http://103.222.20.150/ftpdata/Movies/Indian%20Bangla/2021/"
wshshell.run "LEXI.VBS"

else
if Input = "open a website for downloading latest french movies" OR Input = "Lexi open a website for downloading french latest movies" then
Sapi.speak "Sure here I have a website for downloading french movies"
wshshell.run "http://103.222.20.150/ftpdata/Movies/French_movies/new/"
wshshell.run "LEXI.VBS"

else
if Input = "open a website for downloading latest japanese movies" OR Input = "Lexi open a website for downloading latest japanese movies" then
Sapi.speak "Sure here I have a website for downloading Japanese movies"
wshshell.run "http://103.222.20.150/ftpdata/Movies/Japanese_movies/"
wshshell.run "LEXI.VBS"

else
if Input = "open a website for downloading latest chinese movies" OR Input = "Lexi open a website for downloading latest chinese movies" then
Sapi.speak "Sure here I have a website for downloading chinese movies"
wshshell.run "http://103.222.20.150/ftpdata/Movies/Chines/new/"
wshshell.run "LEXI.VBS"

else
if Input = "open a website for downloading latest english series" OR Input = "Lexi open a website for downloading latest english series" then
Sapi.speak "Sure here I have a website for downloading latest english series from hollywood"
wshshell.run "http://103.222.20.150/ftpdata/Movies/tv%20series/English/"
wshshell.run "LEXI.VBS"

else
if Input = "open a website for downloading latest korean series" OR Input = "Lexi open a website for downloading latest korean series" then
Sapi.speak "Sure here I have a website for downloading latest korean series"
wshshell.run "http://103.222.20.150/ftpdata/Movies/tv%20series/tv%20new/Korean/"
wshshell.run "LEXI.VBS"

else
if Input = "open a website for downloading latest TV series" OR Input = "Lexi open a website for downloading latest Tv series" then
Sapi.speak "Sure here I have a website for downloading latest tv series in every language"
wshshell.run "http://103.222.20.150/ftpdata/Movies/tv%20series/"
wshshell.run "LEXI.VBS"

else
if Input = "open a website for downloading latest movies" OR Input = "Lexi open a website for downloading latest movies" then
Sapi.speak "Sure here I have a website for downloading latest movies in every language"
wshshell.run "http://103.222.20.150/ftpdata/Movies/"
wshshell.run "LEXI.VBS"

else
if Input = "open a website for downloading songs" OR Input = "Lexi open a website for downloading songs" then
Sapi.speak "Sure here I have a website for downloading songs"
wshshell.run "http://f.marvarid.net/mp3/New_mp3/"
wshshell.run "LEXI.VBS"

ELSE
SAPI.SPEAK "SORRY I AM NOT ABLE TO SEARCH "&INPUT
SAPI.SPEAK "SHOULD I SEARCH ABOUT it  ON INTERNET  "
INPUT1=inputbox ("SHOULD I MAKE A GOOGLE SEARCH . PLEASE ANSWER IN YES OR NO")
IF INPUT1="YES" OR INPUT1="yes" THEN
SAPI.SPEAK "I AM MAKING A WEB SEARCH  "
wshshell.run "http://www.google.com/search?q="&INPUT
ELSE
IF INPUT1="NO" OR INPUT1="no" THEN
SAPI.SPEAK " I AM TRYING TO OPEN "&INPUT&" AGAIN IF THE FILE SPECIFIED IS IN YOUR SYSTEM THEN IT WILL OPEN "
SAPI.SPEAK "ELSE PLEASE TRY AGAIN WITH CORRECT NAME "
SAPI.SPEAK " OPENING " & INPUT
wshshell.run INPUT

ELSE
SAPI.SPEAK "SORRY I AM NOT ABLE TO SEARCH "&INPUT
SAPI.SPEAK "SHOULD I SEARCH ABOUT IT ON AMAZON"
INPUT2=inputbox ("SHOULD I MAKE A AMAZON SEARCH . PLEASE ANSWER IN YES OR NO")
IF INPUT2="YES" OR INPUT1="yes" THEN
SAPI.SPEAK "I AM MAKING A AMAZON SEARCH  "
wshshell.run "https://www.amazon.com/s?k="&INPUT
ELSE
IF INPUT2="NO" OR INPUT2="no" THEN
SAPI.SPEAK " I AM TRYING TO OPEN "&INPUT&" AGAIN IF THE FILE SPECIFIED IS IN YOUR SYSTEM THEN IT WILL OPEN "
SAPI.SPEAK "ELSE PLEASE TRY AGAIN WITH CORRECT NAME "
SAPI.SPEAK " OPENING " & INPUT
wshshell.run INPUT

ELSE
SAPI.SPEAK "SORRY I AM NOT ABLE TO SEARCH "&INPUT
SAPI.SPEAK "SHOULD I SEARCH ABOUT IT ON YOUTUBE"
INPUT3=inputbox ("SHOULD I MAKE A YOUTUBE SEARCH. PLEASE ANSWER IN YES OR NO")
IF INPUT3="YES" OR INPUT1="yes" THEN
SAPI.SPEAK "I AM MAKING A YOUTUBE SEARCH  "
wshshell.run "https://www.youtube.com/results?search_query="&INPUT
ELSE
IF INPUT3="NO" OR INPUT3="no" THEN
SAPI.SPEAK " I AM TRYING TO OPEN "&INPUT&" AGAIN IF THE FILE SPECIFIED IS IN YOUR SYSTEM THEN IT WILL OPEN "
SAPI.SPEAK "ELSE PLEASE TRY AGAIN WITH CORRECT NAME "
SAPI.SPEAK " OPENING " & INPUT
wshshell.run INPUT

end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if 
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
