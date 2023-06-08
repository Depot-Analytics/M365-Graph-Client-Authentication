from office365.graph_client import GraphClient
import webbrowser
import msal
import json
import os
import time


def acquire_token_by_web(app, scopes):
   
   # update the last time manual authentication occurred
   last_manual_auth = {"last_manual_auth": time.time()}
   with open("last_manual_auth.json", "w") as f:
       json.dump(last_manual_auth, f)
 
   # authorize through web browser
   flow = app.initiate_device_flow(scopes=scopes)
   # print user code in terminal to copy and paste into web browser
   print("user_code: " + flow["user_code"])
   webbrowser.open(flow["verification_uri"])

   # get token response
   token_response = app.acquire_token_by_device_flow(flow)

   return token_response

def get_access_token_exp(token_information):
    # read access token expiration from token_information.json
    access_token_tag = (
        "3c4e7…"
    )
    access_token_expiration = int(
        token_information["AccessToken"][access_token_tag]["expires_on"]
    )

    return access_token_expiration


def get_account(token_information):
    account_tag = (
        "3c4e7…"
    )
    account_info = token_information["Account"]
    account = {
        "home_account_id": account_info[account_tag]["home_account_id"],
        "environment": account_info[account_tag]["environment"],
        "username": account_info[account_tag]["username"],
        "authority_type": account_info[account_tag]["authority_type"],
        "local_account_id": account_info[account_tag]["local_account_id"],
        "realm": account_info[account_tag]["realm"],
    }

    return account

def get_refresh_token_exp():
    # ensure file exists
    if not os.path.exists("last_manual_auth.json"):
        return 0

    # read time of last manual authentication
    with ("last_manual_auth.json").open("r") as f:
        last_manual_auth = json.load(f)


    # calculate refresh token expiration (90 days after last manual authentication)
    refresh_token_expiration = (
        int(last_manual_auth["last_manual_auth"]) + 7776000
    )

    return refresh_token_expiration


def get_refresh_token(token_information):
    rt_tag = (
        "3c4e7…"
    )
    refresh_token = token_information["RefreshToken"][rt_tag]["secret"]

    return refresh_token


def get_token_response():

    APP_ID = "<enter-your-app-id>"
    SCOPES = ["User.read", "Mail.ReadWrite", "Mail.Send", "Mail.Send.Shared"]

    # create cache and public client app
    cache = msal.SerializableTokenCache()
    app = msal.PublicClientApplication(client_id=APP_ID, token_cache=cache)

    if os.path.exists("token_information.json"):

        # deserialize the cache and read relevant information
        cache.deserialize(open("token_information.json", "r").read())
        with open("token_information.json", "r") as f:
            token_information = json.load(f)

        access_token_expiration = get_access_token_exp(token_information)
        account = get_account(token_information)
        refresh_token_expiration = get_refresh_token_exp()
        refresh_token = get_refresh_token(token_information)

        # get current time
        curr_time = time.time()

        if curr_time < access_token_expiration:
            # access token is still valid
            token_response = app.acquire_token_silent(
                scopes=SCOPES,
                account=account,
            )
        elif curr_time < refresh_token_expiration:
            # access token has expired, but refresh token is still valid
            token_response = app.acquire_token_by_refresh_token(
                scopes=SCOPES, refresh_token=refresh_token
            )
        else:
            # access token has expired, manually authenticate through 		
      # web browser
            token_response = acquire_token_by_web(app, SCOPES)
    else:
        # first time running program - no token info exists, manually 
        # authenticate through web browser
        token_response = acquire_token_by_web(app, SCOPES)

    # write token information to json file
    with open("token_information.json", "w") as f:
        f.write(cache.serialize())

    return token_response


def send_email(subject: str, body: str, to_recipients: list, send_from: str):

    # instantiate client
    client: GraphClient = GraphClient(get_token_response)

    # find desired user
    user = client.users[send_from]

    # send email
    message = user.send_mail(
        subject=subject,
        body=body,
        to_recipients=to_recipients,
    )

    message.execute_query()


send_email(
subject="Testing", 
body="This is my test email!", 
to_recipients=["test.recipient@gmail.com"],  
send_from="test.sender@outlook.com"
)
