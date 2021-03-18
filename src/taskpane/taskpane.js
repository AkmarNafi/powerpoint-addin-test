Office.onReady(info => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("run").onclick = run;

    $("#run").click(run);
  }
});

async function run() {
  getSlideCount();
}

async function getSlideCount() {
  PowerPoint.run(async function(context) {
    context.presentation.load("slides");
    await context.sync();

    var slides = context.presentation.slides;
    var slideCount = slides.getCount();
    await context.sync();

    console.log(slideCount.m_value);
  });
}

async function sendRequest() {
  try {
    //You can change this URL to any web request you want to work with.
    const url = "https://reqres.in/api/users?page=2";
    const response = await fetch(url);
    //Expect that status code is in 200-299 range
    if (!response.ok) {
      throw new Error(response.statusText);
    }
    const jsonResponse = await response.json();

    Office.context.document.setSelectedDataAsync(
      jsonResponse.support.text,
      {
        coercionType: Office.CoercionType.Text
      },
      result => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error(result.error.message);
        }
      }
    );
  } catch (error) {
    return error;
  }
}
