express = require 'express'
http = require 'http'
request = require 'request'
cheerio = require 'cheerio'
graze = require 'graze'
xlsx = require ''
app = express()

PORT = 8000
PORT_TEST = PORT + 1

app.configure ->
  app.set 'port', process.env.PORT or PORT
  app.set 'views', "#{__dirname}/views"
  app.set 'view engine', 'jade'
  app.use express.favicon()
  app.use express.logger('dev') if app.get('env') is 'development'
  app.use express.bodyParser()
  app.use express.methodOverride()
  #app.use express.cookieParser 'your secret here'
  #app.use express.session()
  app.use app.router
  app.use require('connect-assets')(
    helperContext: app.locals
    src: "#{__dirname}/assets"
  )
  app.use express.static "#{__dirname}/public"
  require('./middleware/404')(app)

app.configure 'development', ->
  app.use express.errorHandler()
  app.locals.pretty = true

app.configure "test", ->
  app.set 'port', PORT_TEST

autoload = require('./config/autoload')(app)
autoload "#{__dirname}/helpers", true
autoload "#{__dirname}/assets/js/shared", true
autoload "#{__dirname}/models"
autoload "#{__dirname}/controllers"

require('./config/routes')(app)


url = "http://www.bseindia.com/markets/Equity/newsensexstream.aspx"

template = graze.template {
  '#ctl00_ContentPlaceHolder1_div_iframe':
    results: [
      'iframe':
        url1: graze.attr('src')
    ]
}

template1 = graze.template {
  'td[bgcolor="#d6d6d6"] table tr:not(:first-child)':
    results: [
      '.TTRow_left a':
        Scrip: graze.text()
      '.TTRow:nth-child(3)':
        Price: graze.text()
      '.TTRow:nth-child(6)':
        Buy: graze.text()
      '.TTRow:nth-child(7)':
        Sell: graze.text()
    ]
}

template.scrape(url).then (data) ->
  
  url2 = data.results[0].url1
  console.log url2

  template1.scrape(url2).then (data1) ->
    console.log data1
    xlsx.write 'mySpreadsheet.xlsx', data1.results, (err) ->
      # Error handling here
      return
      
http.createServer(app).listen app.get('port'), ->
  port = app.get 'port'
  env = app.settings.env
  console.log "listening on port #{port} in #{env} mode"

module.exports = app
