/** 
 * Helper function to call MS Graph API endpoint
 * using the authorization bearer token scheme
*/
function callMSGraph(endpoint, token, callback) {
    const headers = new Headers();
    const bearer = `Bearer ${token}`;
    console.log(token);
    headers.append("Authorization", bearer);

    const options = {
        method: "GET",
        headers: headers
    };

    console.log('request made to Graph API at: ' + new Date().toString());

    fetch(endpoint, options)
        .then(response => response.json())
        .then(response => callback(response, endpoint))
        .catch(error => console.log(error));
}

function callDynamicsGraph(endpoint, token, callback){
    if (!token) {
        errorMessage.textContent = 'error occurred: ';
        return;
    }
    
    var req = new XMLHttpRequest()
    req.open("GET", encodeURI("https://customer-support-case-core-test.crm4.dynamics.com/api/data/v9.2/incidents"), true);
    req.setRequestHeader("Authorization", "Bearer " + token);
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.onreadystatechange = function () {
        if (this.readyState == 4 /* complete */) {
            req.onreadystatechange = null;
            if (this.status == 200) {
                var account = JSON.parse(this.response).value;
               // renderAccount(account, token);
            }
            else {
                var error = JSON.parse(this.response).error;
                console.log(error.message);
                errorMessage.textContent = error.message;
            }
        }
    };
    req.send();
}
