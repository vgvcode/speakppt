import os
import boto3

def lambda_handler(event, context):
    #Grab the pipeline ID from the environment variables and the key name from the event passed in
    #The transcoder's bucket is vgvcode.speakppt
    print(os.environ['ELASTIC_TRANSCODER_REGION'])
    print(os.environ['ELASTIC_TRANSCODER_PIPELINE_ID'])
    transcoder = boto3.client('elastictranscoder', os.environ['ELASTIC_TRANSCODER_REGION'])
    pipelineId = os.environ['ELASTIC_TRANSCODER_PIPELINE_ID']
    #example: output/business.txt.pptm.1.mp3
    bucket = event['Records'][0]['s3']['bucket']['name']
    mp3File = event['Records'][0]['s3']['object']['key']

    wavFile = mp3File
    wavFileList = wavFile.split('.')
    wavFileList[0] = wavFileList[0].replace('output/', '')
    wavFileList[-1] = 'wav'
    #Remove the file extension
    wavFile = '.'.join(wavFileList)

    print('Pipe Line Id:' + pipelineId)
    print('Bucket:' + bucket)
    print('Mp3 File:' + mp3File)
    print('Wav File:' + wavFile)

    job = transcoder.create_job(
            PipelineId=pipelineId,
            OutputKeyPrefix='wav/',
            Input={
                'Key': mp3File
            },
            Output={
                    'Key': wavFile,
                    'PresetId': '1351620000001-300300' #Audio WAV 44100 Hz, 16 bit - CD Quality
                }
            )

#const AWS = require('aws-sdk');

# const elasticTranscoder = new AWS.ElasticTranscoder({
#     region: process.env.ELASTIC_TRANSCODER_REGION
# });

#elasticTranscoder = new AWS.ElasticTranscoder({region: os.environ['ELASTIC_TRANSCODER_REGION']})
#const handler = (event, context, callback) => {

#    console.log('Welcome');
#    console.log('event: ' + JSON.stringify(event));

#    // Grab the pipeline ID from the environment variables and the key name from the event passed in
#    const pipelineId = process.env.ELASTIC_TRANSCODER_PIPELINE_ID;
#    const key = event.Records[0].s3.object.key;

#    // The input file may have spaces so replace them with '+'
#    const sourceKey = decodeURIComponent(key.replace(/\+/g, ' '));

#    // Remove the file extension
#    const outputKey = sourceKey.split('.')[0];

#    // Build the parameters for the Job pipeline.
#    // Reference: https://docs.aws.amazon.com/AWSJavaScriptSDK/latest/AWS/ElasticTranscoder.html#createJob-property
#    const params = {
#        PipelineId: pipelineId,
#        OutputKeyPrefix: outputKey + '/',
#        Input: {
#            Key: sourceKey
#        },
#        Outputs: [
#            {
#                Key: outputKey + '-1080p' + '.mp4',
#                PresetId: '1351620000001-000001' //Generic 1080p
#            },
#            {
#                Key: outputKey + '-720p' + '.mp4',
#                PresetId: '1351620000001-000010' //Generic 720p
#            },
#            {
#                Key: outputKey + '-web-720p' + '.mp4',
#                PresetId: '1351620000001-100070' //Web Friendly 720p
#            }
#    ]};
#
#
#     // Call our function that creates an ElasticTranscoder Job
#     return elasticTranscoder.createJob(params)
#         .promise()
#         .then((data) => {
#
#             // Success
#             console.log(`ElasticTranscoder callback data: ${JSON.stringify(data)}`);
#             callback(null, data);
#
#         }).catch((error) => {
#
#             // Failure
#             console.log(`An error occured ${JSON.stringify(error, null, 2)}`);
#             callback(error);
#
#         });
#
# };
#
#
# module.exports = {
#     handler
# };
