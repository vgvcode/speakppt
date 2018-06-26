import sys
import glob
import random
import boto3
from boto3 import client
from subprocess import call
from urllib import parse
import json

s3 = boto3.resource('s3')

#this is the ssml <speak>...</speak>
gCurrentSegment=""
#this is the current slide number
gCurrentSlideNumber=0
#this is a collection of segments; each segment has its own voice id and ssml
# [{VoiceId="", Ssml=""},{...} ...]
gSlideAudioList=[]
#default voice
gDefaultVoice = "Joanna"
#set as default
gCurrentVoice=gDefaultVoice
#this is a collection of voices for this presentation
gVoices = [gDefaultVoice]
#Substitutions for @, # and !
gForwardSubstitutions = {"##" : u'\x81', "@@" : u'\x82', "!!" : u'\x83'}
gReverseSubstitutions = {u'\x81': "#", u'\x82': "@", u'\x83': '!'}

def doVoices(args):
    global gVoices, gCurrentVoice
    gVoices = args
    gCurrentVoice = args[0]

def doGreeting(args):
    if len(args) == 0:
        args = ["Hi", "everyone"]

    greeting = ' '.join(args).split('@')
    addToSegment('<prosody rate="slow" volume="x-loud">')
    for g in greeting:
        if len(g) > 0:
            addToSegment('<s>' + g + '</s>')
    addToSegment("</prosody>")

def doGreetingInteractive(args):
    addToSegment("<s>Hi everyone!</s>")
    addToSegment("<s>How are you doing?</s>")
    addToSegment('<break time="2s"/>')
    addToSegment('<prosody volume="x-loud">Louder, please. I can hardly hear you</prosody>')
    addToSegment('<break time="2s"/>')
    addToSegment('<prosody volume="x-loud">OK, much better, thank you!</prosody>')

def doSmallTalk(args):
    for a in args:
        if a == "weekend":
            addToSegment("<s>So, how was your weekend?</s>")
            addToSegment("<s>Good?</s>")
        elif a == "weather":
            addToSegment("<s>The weather has been crazy this week</s>")
            addToSegment("<s>I hope it gets better soon</s>")
        elif a == "traffic":
            addToSegment("<s>The traffic was really bad getting in this morning</s>")
            addToSegment("<s>I hope you all had a better commute</s>")
        else:
            pass

def doIntroduceMe(args):
    addToSegment("<s>My name is " + gCurrentVoice + "</s>")
    addToSegment('<break time="1s"/>')

def doTitle(args):
    title = ' '.join(args).split('@')
    addToSegment('<prosody volume="x-loud">Today we are going to talk about ')
    for t in title:
        if len(t) > 0:
            addToSegment('<s>' + t + '</s>')
    addToSegment("</prosody>")

def doPause(args):
    addToSegment('<break time="' + args[0] + 's"/>')

def doParagraph(args):
    para = ' '.join(args).split('@')
    addToSegment('<break time="1s"/>')
    for p in para:
        if len(p) > 0:
            addToSegment('<p>' + p + '</p>')

def doHeading(args):
    heading = ' '.join(args).split('@')
    addToSegment('<prosody volume="x-loud">')
    for h in heading:
        if len(h) > 0:
            addToSegment('<s>' + h + '</s>')
    addToSegment('</prosody>')

def doBullets(args):
    bullets = ' '.join(args).split("@")
    num = 1
    addToSegment('<break time="1s"/>')
    ix = 0
    for b in bullets:
        ix = ix + 1
        if len(b) > 0:
            if ix == len(bullets):
                addToSegment('<s>And finally, ' + str(num) + "</s><s>" + b + '</s>')
            else:
                addToSegment('<s>' + str(num) + "</s><s>" + b + '</s>')
            num = num + 1

def doQuiz(args):
    question = ' '.join(args).split('@')
    addToSegment('<prosody volume="x-loud">OK, time for a short quiz')
    for q in question:
        if len(q) > 0:
            addToSegment('<s>' + q + '</s>')
    addToSegment('</prosody>')

def doStatusCheck(args):
    addToSegment('<break time="2s"/>')
    addToSegment('<prosody volume="x-loud">')
    addToSegment('<s>OK, Is everyone with me so far?</s>')
    addToSegment('</prosody>')

def doTimeCheck(args):
    addToSegment('<break time="2s"/>')
    addToSegment('<prosody volume="x-loud">')
    addToSegment('<s>OK, let''s do a quick time check</s>')
    addToSegment('<s>How are we doing with the time?</s>')
    addToSegment('</prosody>')

def doQuestions(args):
    addToSegment('<prosody volume="x-loud">')
    addToSegment('<s>OK, that was a lot of material.</s>')
    addToSegment('<s>Are there any questions?</s>')
    addToSegment('</prosody>')

def doTransition(args):
    global gCurrentSlideNumber, gCurrentSegment, gCurrentVoice
    try:
        if gCurrentSlideNumber == 0:
            #create a new segment
            gCurrentSlideNumber = 1
        else:
            #close out previous segment
            gCurrentSegment = gCurrentSegment + "</speak>"
            addToSlideAudioList(gCurrentVoice, gCurrentSegment, gSlideAudioList)
            #increment the slide number and initialize next segment
            gCurrentSlideNumber = gCurrentSlideNumber + 1
            gCurrentVoice = getNextVoice(gCurrentVoice, gVoices)

        #if a voice is specified in the transition command, honor it
        if len(args) > 0:
            gCurrentVoice = args[0]

        #open a new segment, assign voice
        gCurrentSegment = "<speak>"
        addToSegment('<break time="2s"/>')
    except Exception:
        print("Unexpected error @doTransition:", sys.exc_info()[0])

def doSummary(args):
    summary = ' '.join(args).split('@')
    addToSegment('<s><prosody volume="x-loud">So, let me summarize what I just said</prosody></s>')
    for s in summary:
        if len(s) > 0:
            addToSegment('<s>' + s + '</s>')

def doConclusion(args):
    conclusion = ' '.join(args).split('@')
    addToSegment('<prosody volume="x-loud"><s>That''s all folks.</s>')
    addToSegment('<s>I hope you enjoyed this presentation.</s>')
    addToSegment('<s>Thank you for listening</s></prosody>')
    for c in conclusion:
        if len(c) > 0:
            addToSegment('<s>' + c + '</s>')
    #do a final transition here to add this segment to the running audio list
    doTransition("")

def processTag(section):
    try:
        words = section.split(' ')
        tag = words[0]
        args = words[1:]

        switcher = {
            'voices': lambda: doVoices(args),
            'greeting': lambda: doGreeting(args),
            'greeting-interactive': lambda: doGreetingInteractive(args),
            'smalltalk': lambda: doSmallTalk(args),
            'introduce-me': lambda: doIntroduceMe(args),
            'title': lambda: doTitle(args),
            'pause': lambda: doPause(args),
            'paragraph': lambda: doParagraph(args),
            'heading': lambda: doHeading(args),
            'bullets': lambda: doBullets(args),
            'transition': lambda: doTransition(args),
            'quiz': lambda: doQuiz(args),
            'questions': lambda: doQuestions(args),
            'status-check': lambda: doStatusCheck(args),
            'time-check': lambda: doTimeCheck(args),
            'summary': lambda: doSummary(args),
            'conclusion': lambda: doConclusion(args)
        }

        #print('Tag:' + tag)
        #print(args)
        func = switcher.get(tag, lambda: '')
        return func()
    except Exception:
        print("Unexpected error @doProcessTag:", sys.exc_info()[0])
        return None

def addToSegment(s):
    global gCurrentSegment
    args = []
    args.append(s)
    args = doReverseSubstitutions(args)
    s = ''.join(args)
    gCurrentSegment = gCurrentSegment + " " + s

def addToSlideAudioList(voice, segment, audioList):
    audioList.append({"VoiceId": voice, "Ssml": segment})

def getNextVoice(currentVoice, listOfVoices):
    try:
        if gCurrentVoice in gVoices:
            ix = gVoices.index(gCurrentVoice)
            ix = (ix + 1) % len(gVoices)
            nextVoice = gVoices[ix]
        else:
            nextVoice = gDefaultVoice
    except Exception:
        #no voice found in the list, setting to default voice
        print("Unexpected error @getNextVoice:", sys.exc_info()[0])
        nextVoice = gDefaultVoice
    finally:
        return nextVoice

def convertMp3ToWav(baseName, segmentId):
    #baseName = business.txt (or) /tmp/business.txt
    mp3 = baseName + ".pptm." + str(segmentId) + ".mp3"
    wav = baseName + ".pptm." + str(segmentId) + ".wav"
    #mp3 = <baseName>.pptm.1.mp3
    #wav = <baseName>.pptm.1.wav
    #convert mp3 to wav
    call(["ffmpeg", "-i", mp3, wav])
    #remove the mp3
    call(["rm", mp3])
    print('Converted ' + mp3 + ' to ' + wav)

def saveSlideAudioList(audioList, baseFileName):
    #baseFileName = business.txt (or) /tmp/business.txt
    polly = client("polly", region_name="us-east-1")
    segmentId = 0
    #write out one mp3 file for each slide
    for a in audioList:
        segmentId = segmentId + 1
        response = polly.synthesize_speech(Text=a['Ssml'], TextType='ssml', OutputFormat="mp3", VoiceId=a['VoiceId'])
        data_stream=response.get('AudioStream')
        stream=data_stream.read()
        outputMp3FileName = baseFileName + ".pptm." + str(segmentId) + ".mp3"
        #outputMp3FileName = <baseFileName>.pptm.1.mp3
        f=open(outputMp3FileName, "wb")
        f.write(stream)
        f.close()
        print('Created ' + outputMp3FileName)
        #convertMp3ToWav(baseFileName, segmentId)

def doForwardSubstitutions(s):
    global gForwardSubstitutions
    for key in gForwardSubstitutions:
        s = s.replace(key, gForwardSubstitutions[key])
    return s

def doReverseSubstitutions(args):
    global gReverseSubstitutions
    newArgs = []
    for a in args:
        for key in gReverseSubstitutions:
            a = a.replace(key, gReverseSubstitutions[key])
        newArgs.append(a)
    return newArgs

def doTranslate(fileName):
    #filename = business.txt
    f=open(fileName, "r")
    lines=f.readlines()
    txt=''.join(lines).replace('\n', ' ')
    txt=doForwardSubstitutions(txt)
    sections=txt.split("#")
    #print(sections)
    for s in sections:
        s = s.strip()
        if len(s) > 0:
            #print('Processing:' + s)
            processTag(s)

    # for a in gSlideAudioList:
    #     print('Speaker:' + a['VoiceId'])
    #     print('Speech:' + a['Ssml'])
    #     print

    saveSlideAudioList(gSlideAudioList, fileName)

def doTranslateInLambda(b, s3Inf):
    #b = vgvcode.speakppt
    #s3Inf = input/business.txt
    global gPrs

    #read the data file
    data = readFromS3Bucket(b, s3Inf).decode('utf-8')
    print("Read " + s3Inf + " from s3")

    lines = data.split('\n')
    txt=''.join(lines).replace('\n', ' ')
    txt=doForwardSubstitutions(txt)
    sections=txt.split("#")
    #print(sections)
    for s in sections:
        s = s.strip()
        if len(s) > 0:
            #print('Processing:' + s)
            processTag(s)

    #create base file name
    baseName = s3Inf.replace("input", "/tmp", 1)
    #tmpOutf = /tmp/business.txt
    saveSlideAudioList(gSlideAudioList, baseName)
    moveAudioFilesToS3(b)

# read file from s3
def readFromS3Bucket(b, s3Inf):
    #return bytes
    return s3.Object(b, s3Inf).get()['Body'].read()

def moveAudioFilesToS3(b):
    #b = vgvcode.speakppt
    #path = "/tmp/*.wav"
    path = "/tmp/*.mp3"
    fileList = glob.glob(path)
    for f in fileList:
        #f = /tmp/business.txt.1.pptm.wav
        s3Outf = f.replace("/tmp", "output", 1)
        s3.Object(b, s3Outf).put(Body=open(f, 'rb'))
        print("Copied file to " + s3Outf)

def lambda_handler(event, context):
    #bucket = vgvcode.speakppt
    #dataFile = input/business.txt
    print(json.dumps(event))
    snsMsgObject = json.loads(event['Records'][0]['Sns']['Message'])
    bucket = snsMsgObject['Records'][0]['s3']['bucket']['name']
    key = snsMsgObject['Records'][0]['s3']['object']['key']

    #Object key may have spaces or unicode non-ASCII characters.
    dataFile = parse.unquote(key.replace('/\+/g', " "))


    snsMsgObject = json.loads(event['Records'][0]['Sns']['Message'])
    bucket = snsMsgObject['Records'][0]['s3']['bucket']['name']
    key = snsMsgObject['Records'][0]['s3']['object']['key']

    #Object key may have spaces or unicode non-ASCII characters.
    dataFile = parse.unquote(key.replace('/\+/g', " "))
    print("Bucket:" + bucket + ", dataFile:" + dataFile)
    doTranslateInLambda(bucket, dataFile)

def main():
    if len(sys.argv) != 2:
        print("Usage:" + sys.argv[0] + " /path/to/manifest")
        exit()
    fileName = sys.argv[1]
    #filename = business.txt
    doTranslate(fileName)

if __name__ == "__main__":
    # execute only if run as a script
    main()
