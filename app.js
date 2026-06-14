const tabButtons =
    document.querySelectorAll(".tab-button");

const tabContents =
    document.querySelectorAll(".tab-content");

tabButtons.forEach(button => {

    button.addEventListener("click", () => {

        const tab =
            button.dataset.tab;

        // remove active
        tabButtons.forEach(btn => {
            btn.classList.remove("active");
        });

        tabContents.forEach(content => {
            content.classList.remove("active");
        });

        // activate current
        button.classList.add("active");

        document
            .getElementById(`${tab}-tab`)
            .classList.add("active");

    });

});
