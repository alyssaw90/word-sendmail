# word-sendmail

To run the sample add-in code under node: 

1. git clone https://github.com/tomjebo/word-sendmail.git

2. cd word-sendmail\AdatumWordWeb

3. change the redirect URL in the app registration to port 3000

4. change the following code in home.js to use port 3000 instead of 44300:

~~~
        authenticator.endpoints.registerMicrosoftAuth(clientId, {
            redirectUrl: 'https://localhost:3000/home.html'/* , scope: 'User.Read.All Mail.Send'*/
        });
~~~

5. open the manifest and change all ~remoteAppUrl to http://localhost:3000 like this example:

~~~
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:3000/Home.html" />
  </DefaultSettings>
~~~

6. npm install

7. npm start

You should see something like this output:

~~~
$ npm start

> AdatumWord@0.1.0 start C:\Users\tomjebo\Source\Repos\word-sendmail\AdatumWordWeb
> browser-sync start --config bsconfig.json

[BS] Access URLs:
 ----------------------------------------
       Local: https://localhost:3000
    External: https://172.19.188.212:3000
 ----------------------------------------
          UI: http://localhost:3001
 UI External: http://172.19.188.212:3001
 ----------------------------------------
[BS] Serving files from: ./
[BS] Watching files...
~~~


