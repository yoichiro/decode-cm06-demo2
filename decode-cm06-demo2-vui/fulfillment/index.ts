import { actionssdk, SignIn, Suggestions, ActionsSdkConversation } from 'actions-on-google'
import * as express from 'express'
import * as request from 'request';
import * as moment from 'moment'
const { createHandler } = require('azure-function-express')

const app = actionssdk({
  clientId: process.env.CLIENT_ID
})

const rooms = {
  'RoomAdams': {
    email: 'Adams@M365x850511.onmicrosoft.com', name: 'Conf Room Adams'
  },
  'RoomBaker': {
    email: 'Baker@M365x850511.onmicrosoft.com', name: 'Conf Room Baker'
  },
  'RoomCrystal': {
    email: 'Crystal@M365x850511.onmicrosoft.com', name: 'Conf Room Crystal'
  },
  'RoomHood': {
    email: 'Hood@M365x850511.onmicrosoft.com', name: 'Conf Room Hood'
  },    
  'RoomRainer': {
    email: 'Rainier@M365x850511.onmicrosoft.com', name: 'Conf Room Rainier'
  },
  'RoomStevens': {
    email: 'Stevens@M365x850511.onmicrosoft.com', name: 'Conf Room Stevens'
  }
}

app.intent('actions.intent.MAIN', conv => {
  const accessToken = conv.user.access.token
  if (accessToken) {
    conv.ask('こんにちは!最近の予定を知りたいですか？それとも、会議予定を登録したいですか？')
    conv.ask(new Suggestions(['最近の予定', '会議の予約']))
  } else {
    conv.ask(new SignIn())
  }
})

app.intent('actions.intent.SIGN_IN', (conv, params, signin) => {
  if (signin['status'] === 'OK') {
    conv.ask('こんにちは!最近の予定を知りたいですか？それとも、会議予定を登録したいですか？')
    conv.ask(new Suggestions(['最近の予定', '会議の予約']))
  } else {
    conv.close('アカウントリンクが必要です。')
  }
})

app.intent('actions.intent.TEXT', async (conv, raw) => {
  const context = conv.data['context']
  if (!context) {
    const intent = await decideIntent(conv, raw)
    if (intent.topScoringIntent.intent === 'recent-events') {
      return await handleRecentEventsIntent(conv, raw)
    } else if (intent.topScoringIntent.intent === 'create-event') {
      return handleCreateEventIntent(conv, raw)
    }
  } else if (context === 'ask-event-name') {
    return handleDecideEventNameIntent(conv, raw)
  } else if (context === 'ask-event-date') {
    const intent = await decideIntent(conv, raw)
    if (intent.topScoringIntent.intent === 'event-date') {
      return handleDecideEventDateIntent(conv, raw, intent)
    }
  } else if (context === 'ask-event-time') {
    const intent = await decideIntent(conv, raw)
    if (intent.topScoringIntent.intent === 'event-time') {
      return handleDecideEventTimeIntent(conv, raw, intent)
    }
  } else if (context === 'ask-event-room') {
    const intent = await decideIntent(conv, raw)
    if (intent.topScoringIntent.intent === 'event-room') {
      return handleDecideEventRoomIntent(conv, raw, intent)
    }
  } else if (context === 'confirm-create-event') {
    const intent = await decideIntent(conv, raw)
    if (intent.topScoringIntent.intent === 'confirm-yes') {
      return await handleConfirmYesIntent(conv, raw, intent)
    }
  }
  conv.ask('よくわかりませんでした。もう一度おっしゃってください。')
})

const handleConfirmYesIntent = (conv: ActionsSdkConversation, raw: string, intent: any): Promise<void> => {
  return new Promise((resolve, reject) => {
    request({
      url: process.env.BUSINESS_LOGIC_URL,
      method: 'POST',
      headers: {
        Authorization: `Bearer ${conv.user.access.token}`
      },
      json: {
        subject: conv.data['eventName'],
        startDate: conv.data['eventDate'],
        startTime: conv.data['eventTime'],
        length: conv.data['eventLength'],
        location: conv.data['eventRoom']
      }
    }, (error, response, body) => {
      if (error) {
        reject(error)
      } else {
        conv.close(`会議登録を行いました。また会いましょう。`)
        resolve()
      }
    })
  })
}

const handleDecideEventRoomIntent = (conv: ActionsSdkConversation, raw: string, intent: any): void => {
  const roomEntity = intent.entities[0]
  const room = rooms[roomEntity.type]
  conv.data['eventRoom'] = `${room.email}/${room.name}`
  conv.ask(`ありがとうございます。では、${conv.data['eventName']}の会議予約を行いますか？`)
  conv.data['context'] = 'confirm-create-event'
}

const handleDecideEventTimeIntent = (conv: ActionsSdkConversation, raw: string, intent: any): void => {
  intent.entities.forEach((entity: any) => {
    if (entity.type === 'RegexpTime') {
      const m = entity.entity.match(/([0-9]+)時([0-9]+)分/)
      conv.data['eventTime'] = `${m[1]}:${m[2]}`
    } else if (entity.type === 'RegexpHour') {
      const m = entity.entity.match(/([0-9]+)時/)
      conv.data['eventTime'] = `${m[1]}:00`
    } else if (entity.type === 'RegexpLength') {
      const m = entity.entity.match(/([0-9]+)時間/)
      conv.data['eventLength'] = String(Number(m[1]) * 60)
    }
  });
  conv.ask('どの会議室を予約しますか？')
  conv.data['context'] = 'ask-event-room'
}

const handleDecideEventDateIntent = (conv: ActionsSdkConversation, raw: string, intent: any): void => {
  const entity = intent.entities[0]
  if (entity.type === 'RegexpDate') {
    const m = entity.entity.match(/([0-9]+)年([0-9]+)月([0-9]+)日/)
    conv.data['eventDate'] = `${m[1]}/${m[2]}/${m[3]}`
  } else if (entity.type === 'DateExpression') {
    if (entity.entity === '明日' || entity.entity === 'あした') {
      conv.data['eventDate'] = moment().add('d', 1).format('YYYY/MM/DD')
    } else if (entity.entity === '今日' || entity.entity === 'きょう') {
      conv.data['eventDate'] = moment().format('YYYY/MM/DD')
    } else if (entity.entity === '明後日' || entity.entity === 'あさって') {
      conv.data['eventDate'] = moment().add('d', 2).format('YYYY/MM/DD')
    }
  }
  conv.ask('何時何分から何時間行いますか？')
  conv.data['context'] = 'ask-event-time'
}

const handleDecideEventNameIntent = (conv: ActionsSdkConversation, raw: string): void => {
  conv.data['eventName'] = raw
  conv.ask('何年何月何日に行いますか？')
  conv.data['context'] = 'ask-event-date'
}

const handleCreateEventIntent = (conv: ActionsSdkConversation, raw: string): void => {
  conv.ask('会議の名前はなんですか？')
  conv.data['context'] = 'ask-event-name'
}

const handleRecentEventsIntent = (conv: ActionsSdkConversation, raw: string): Promise<void> => {
  return new Promise((resolve, reject) => {
    request({
      url: process.env.BUSINESS_LOGIC_URL,
      headers: {
        Authorization: `Bearer ${conv.user.access.token}`
      },
      json: true
    }, (error, response, body) => {
      if (error) {
        reject(error)
      } else {
        if (body.value.length === 0) {
          conv.close(`今後7日間に予定はありません。`)
        } else {
          conv.close(`今後7日間に予定は${body.value.length}件あります。直近の予定は、${body.value[0].subject}です。`)
        }
        resolve()
      }
    })
  })
}

const decideIntent = (conv: ActionsSdkConversation, query: string): Promise<any> => {
  return new Promise((resolve, reject) => {
    request({
      url: process.env.LUIS_URL,
      qs: {
        verbose: 'true',
        timezoneOffset: '-360',
        'subscription-key': process.env.LUIS_SUBSCRIPTION_KEY,
        q: query
      },
      json: true
    }, (error, response, body) => {
      if (error) {
        reject(error)
      } else {
        resolve(body)
      }
    })
  })
}

const expressApp = express()
expressApp.post('/api/fulfillment', app)

module.exports = createHandler(expressApp)
