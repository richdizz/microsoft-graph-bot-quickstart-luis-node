import { IDialog } from './idialog';
import * as builder from 'botbuilder';
import * as restify from 'restify';
import { AuthHelper } from '../helpers/authHelper';
import { HttpHelper } from '../helpers/httpHelper';
import { SentimentHelper } from '../helpers/sentimentHelper';

export class searchFilesDialog implements IDialog {
    constructor(private authHelper: AuthHelper) {
        this.id = 'SearchFiles';
        this.name = 'SearchFiles';
        this.waterfall = [].concat(
            (session, args, next) => {
                // Read the LUIS detail and then move to auth
                this.keywords = encodeURIComponent(this.getIntentKeywords(args));
                next();                
            },
            authHelper.getAccessToken(),
            (session, results, next) => {
                if (results.response != null) {
                    // make a call to the Microsoft Graph to search files
                    let headers = {
                        Accept: 'application/json',
                        Authorization: 'Bearer ' + results.response
                    };
                    let endpoint = `/v1.0/me/drive/search(q='${this.keywords}')?$select=id,name,size,webUrl&$top=5`;
                    HttpHelper.getJson(headers, 'graph.microsoft.com', endpoint).then(function(data: any) {
                        // send the results as a carousel
                        var msg = new builder.Message(session);
                        msg.attachmentLayout(builder.AttachmentLayout.carousel);
                        let cards = [];
                        for (var i = 0; i < data.value.length; i++) {
                            cards.push(
                                new builder.HeroCard(session)
                                    .title(data.value[i].name)
                                    .subtitle(`Size: ${data.value[i].size}`)
                                    .text(`Download: ${data.value[i].webUrl}`)
                            );
                        }
                        msg.attachments(cards);
                        session.send(msg).endDialog();
                    }).catch(function(err) {
                        // something went wrong
                        session.endConversation(`Error occurred: ${err}`);
                    });
                }
                else {
                    session.endConversation('Sorry, I did not understand');
                }
            }
        );
    }
    
    id; name; waterfall; keywords;

    getIntentKeywords(args: any) {
        if (args.intent.entities.length == 0) return null;
        else {
            var entity = args.intent.entities[0];
            if (entity.type == 'FileType') {
                // perform search based on filetype...but clean up the filetype first
                let fileType = entity.entity.replace(' . ', '.').replace('. ', '.').toLowerCase();
                let images: Array<string> = [ 'images', 'pictures', 'pics', 'photos', 'image', 'picture', 'pic', 'photo' ];
                let presentations: Array<string>  = [ 'powerpoints', 'presentations', 'decks', 'powerpoints', 'presentation', 'deck', '.pptx', '.ppt', 'pptx', 'ppt' ];
                let documents: Array<string> = [ 'documents', 'document', 'word', 'doc', '.docx', '.doc', 'docx', 'doc' ];
                let workbooks: Array<string> = [ 'workbooks', 'workbook', 'excel', 'spreadsheet', 'spreadsheets', '.xlsx', '.xls', 'xlsx', 'xls' ];
                let music: Array<string> = [ 'music', 'songs', 'albums', '.mp3', 'mp3' ];
                let videos: Array<string> = [ 'video', 'videos', 'movie', 'movies', '.mp4', 'mp4', '.mov', 'mov', '.avi', 'avi' ];

                if (images.indexOf(fileType) != -1)
                    return '.png .jpg .jpeg .gif';
                else if (presentations.indexOf(fileType) != -1)
                    return '.pptx .ppt';
                else if (documents.indexOf(fileType) != -1)
                    return '.docx .doc';
                else if (workbooks.indexOf(fileType) != -1)
                    return '.xlsx .xls';
                else if (music.indexOf(fileType) != -1)
                    return '.mp3 .wav';
                else if (videos.indexOf(fileType) != -1)
                    return '.mp4 .avi .mov';
                else
                    return fileType;
            }
            else if (entity.type == 'FileName') {
                return entity.entity;
            }
        }
    }
}

export default searchFilesDialog;