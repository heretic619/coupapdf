const express = require('express');
const bodyparser = require('body-parser');

var AWS = require('aws-sdk');
var fs = require('fs');
var rimraf = require('rimraf');

AWS.config.update({
    accessKeyId: 'AKIAJOSOBH4IVM5CP77A', 
    secretAccessKey: '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'
})

var app = express();
app.use(bodyparser.json());

var PythonShell = require('python-shell');

app.post('/generate', (req, res) => {
    var body = req.body;
    var options = {
        mode: 'text',
        pythonPath: 'python3',
        args: [
            `-process=${body.process}`,
            `-username=${body.username}`,
            `-password=${body.password}`,
            `-site=${body.site}`,
            `-path=${body.path}`,
            `-csvpath=${body.csvpath}`,
            `-stylesheet=${body.stylesheet}`,
            `-outfolder=${body.outfolder}`,
            `-email=benr@mindtouch.com`,
            `-title_long=${body.title_long}`,
            `-title_short=${body.title_short}`,
            `-doc_status=${body.doc_status}`,
            `-copyright=${body.copyright}`,
            `-legal_message=${body.legal_message}`,
            `-valid_date=${body.valid_date}`
        ]
    };

    PythonShell.run('pdfexporter.py', options, function (err, results) {
        if (err) throw err;
        if (results[0].includes("COMPLETE")) {
            console.log(results[0].split(' '));
            var path = results[0].split(' ')[2];
            var pdf = fs.readFileSync('./' + path + '/CombinedPagesFinalbookmarked.pdf');

            var now = new Date();
            var filename = "CoupaFinalPDF_" + now.toISOString().slice(0,10).replace(/-/g,"") + ".pdf";

            var base64data = new Buffer(pdf, 'binary');
            var s3 = new AWS.S3();
            s3.upload({
                Bucket: 'coupa-pdf',
                Key: filename,
                Body: base64data,
                ACL: 'public-read'
            }, function(resp) {
                console.log(arguments);

                if (arguments[1]['Location']) {

                    var body = `Download the PDF at ${arguments[1]['Location']}`;

                    var params = {
                        Destination: { /* required */
                          ToAddresses: [
                            'benr@mindtouch.com',
                            /* more items */
                          ]
                        },
                        Message: { /* required */
                          Body: { /* required */
                            Html: {
                             Charset: "UTF-8",
                             Data: body
                            }
                           },
                           Subject: {
                            Charset: 'UTF-8',
                            Data: 'Test email'
                           }
                          },
                        Source: 'benr@mindtouch.com', /* required */
                        ReplyToAddresses: [
                            'benr@mindtouch.com'
                        ],
                      };     
                      
                      AWS.config.update({region: 'us-east-1'});
                      
                      // Create the promise and SES service object
                      var sendPromise = new AWS.SES({apiVersion: '2010-12-01'}).sendEmail(params).promise();
                      
                      // Handle promise's fulfilled/rejected states
                      sendPromise.then(
                        function(data) {
                          console.log(data.MessageId);
                        }).catch(
                          function(err) {
                          console.error(err, err.stack);
                        });
                }

                rimraf.sync('./' + path.split('/')[0]);
            });
        }
    });

    res.send('Complete');
});

app.get('/', (req, res) => {
    console.log('GET');
    res.send();
});

app.listen(3000, ()=> {
    console.log("App listening on port 3000");
});