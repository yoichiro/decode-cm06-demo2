import * as express from 'express'
import * as url from 'url'
import * as request from 'request'
import * as cookieParser from 'cookie-parser'
import * as moment from 'moment'
import * as bodyParser from 'body-parser'

const app = express()
app.set('view engine', 'ejs')
app.use(cookieParser())
app.use(bodyParser.urlencoded({extended: true}))

const TENANT_ID = process.env.TENANT_ID
const AUTH_BASE_URL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0`
const CLIENT_ID = process.env.CLIENT_ID
const CLIENT_SECRET = process.env.CLIENT_SECRET

app.get('/', (req, res) => {
    const accessToken = req.cookies.accessToken
    if (accessToken) {
        renderForm(req, res)
    } else {
        renderLoginPage(req, res)
    }
})

const renderForm = (req: express.Request, res: express.Response): void => {
    const accessToken = req.cookies.accessToken
    request({
        url: process.env.BUSINESS_LOGIC_URL!,
        method: 'GET',
        json: true,
        headers: {
            'Authorization': `Bearer ${accessToken}`
        }
    }, (error, response, body) => {
        if (error) {
            res.status(500).send(error)
        } else {
            res.render('form.ejs', {
                events: body.value.map(event => {
                    return {
                        start: moment(event.start.dateTime).format('YYYY/MM/DD HH:mm'),
                        subject: event.subject,
                        location: event.location.displayName
                    }
                })
            })
        }
    })
}

const renderLoginPage = (req: express.Request, res: express.Response): void => {
    const endpointUrl = `${AUTH_BASE_URL}/authorize`
    const queryParams = {
        'client_id': CLIENT_ID,
        'response_type': 'code',
        'redirect_uri': `${process.env.APP_URL}/callback`,
        'scope': 'User.Read Calendars.ReadWrite.Shared',
        'prompt': 'consent'
    }
    const targetUrl = url.parse(endpointUrl, true)
    const query = targetUrl.query
    Object.keys(queryParams).forEach(key => {
        query[key] = queryParams[key]
    })
    res.render('login.ejs', {
        authorizationUrl: url.format(targetUrl)
    })
}

app.get('/callback', (req, res) => {
    request({
        url: `${AUTH_BASE_URL}/token`,
        method: 'POST',
        form: {
            'grant_type': 'authorization_code',
            'client_id': CLIENT_ID,
            'code': req.query.code,
            'redirect_uri': `${process.env.APP_URL}/callback`,
            'client_secret': CLIENT_SECRET
        },
        json: true
    }, (error, response, body) => {
        if (error) {
            res.status(500).send(error)
        } else {
            res.cookie('accessToken', body.access_token, {
                maxAge: 60 * 60 * 1000,
                httpOnly: false
            })
            res.redirect('/')
        }
    })
})

app.post('/', (req, res) => {
    const accessToken = req.cookies.accessToken
    request({
        url: process.env.BUSINESS_LOGIC_URL!,
        method: 'POST',
        json: {
            subject: req.body.subject,
            location: req.body.location,
            startDate: req.body.startDate,
            startTime: req.body.startTime,
            length: req.body.length
        },
        headers: {
            'Authorization': `Bearer ${accessToken}`
        }
    }, (error, response, body) => {
        if (error) {
            res.status(500).send(error)
        } else {
            res.redirect('/')
        }
    })
})

app.listen(process.env.PORT || 1337)
