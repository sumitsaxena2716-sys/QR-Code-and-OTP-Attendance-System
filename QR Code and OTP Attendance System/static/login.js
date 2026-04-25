function login(){

    let username = document.getElementById("username").value;
    let password = document.getElementById("password").value;

    fetch("/check-login", {
        method: "POST",
        headers: {
            "Content-Type": "application/json"
        },
        body: JSON.stringify({
            username: username,
            password: password
        })
    })
    .then(res => res.json())
    .then(data => {

        if(data.status === "ok"){
            window.location.href = "/dashboard-home";
        } else {
            alert("❌ Invalid Username or Password");
        }

    })
    .catch(err => {
        console.log(err);
        alert("Server error");
    });
}