import { AzureFunction, Context, HttpRequest } from '@azure/functions'
import * as request from 'request'
import * as moment from 'moment'

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    if (req.method == 'GET') {
        context.res = await getRecentEvents(context, req)
    } else if (req.method == 'POST') {
        context.res = await createEvent(context, req)
    } else {
        throw new Error('Method not allowed')
    }
}

const getRecentEvents = function(context: Context, req: HttpRequest): Promise<object> {
    const accessToken = req.headers.authorization.split(' ')[1]
    return new Promise((resolve, reject) => {
        request({
            url: 'https://graph.microsoft.com/v1.0/me/calendar/calendarView',
            qs: {
                'startDateTime': new Date().toISOString(),
                'endDateTime': new Date(Date.now() + 604800000).toISOString(),
                '$orderby': 'start/datetime'
            },
            method: 'GET',
            json: true,
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Prefer': 'outlook.timezone="Asia/Tokyo"'
            }
        }, (error, response, body) => {
            if (error) {
                resolve({
                    status: 500,
                    body: error,
                    headers: {
                        'Content-Type': 'application/json'
                    }
                })
            } else {
                resolve({
                    status: 200,
                    body,
                    headers: {
                        'Content-Type': 'application/json'
                    }
                })
            }
        })
    })
}

const createEvent = function(context: Context, req: HttpRequest): Promise<object> {
    const params = req.body
    const accessToken = req.headers.authorization.split(' ')[1]
    return new Promise((resolve, reject) => {
        const subject = params.subject
        const [locationEmailAddress, locationName] = params.location.split('/')
        const startDate = params.startDate.split('/')
        const startTime = params.startTime.split(':')
        const startDateTime = moment([
            Number(startDate[0]), Number(startDate[1]) - 1, Number(startDate[2]),
            Number(startTime[0]), Number(startTime[1]), 0])
        const length = params.length
        const endDateTime = moment(startDateTime).add(length, 'm')
        request({
            url: 'https://graph.microsoft.com/v1.0/me/calendar/events',
            method: 'POST',
            json: {
                subject,
                start: {
                    dateTime: startDateTime.format('YYYY-MM-DDTHH:mm:ss'),
                    timeZone: 'Asia/Tokyo'
                },
                end: {
                    dateTime: endDateTime.format('YYYY-MM-DDTHH:mm:ss'),
                    timeZone: 'Asia/Tokyo'
                },
                location: {
                    locationEmailAddress,
                    displayName: locationName
                },
                attendees: [
                    {
                        emailAddress: {
                            address: locationEmailAddress
                        },
                        type: 'resource'
                    }
                ]
            },
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Prefer': 'outlook.timezone="Asia/Tokyo"'
            }
        }, (error, response, body) => {
            if (error) {
                resolve({
                    status: 500,
                    body: error,
                    headers: {
                        'Content-Type': 'application/json'
                    }
                })
            } else {
                resolve({
                    status: 200,
                    body,
                    headers: {
                        'Content-Type': 'application/json'
                    }
                })
            }
        })
    })
}

export default httpTrigger
