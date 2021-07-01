Office.onReady(info => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("run").onclick = run;
  }
});

function updateStatus(message) {
  console.log(message);
}

async function run() {
  // sendFile();
  getSelectionData();
}

function getSelectionData() {
  PowerPoint.run(function(context) {
    console.log(context);

    var selection = context.presentation.getSelection();

    selection.font.bold = true;

    return context.sync().then(function() {
      console.log("The selection is now bold.");
    });
  }).catch(function(error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

// Get all of the content from a PowerPoint or Word document in 100-KB chunks of text.
function sendFile() {
  Office.context.document.getFileAsync("compressed", { sliceSize: 100000 }, function(result) {
    if (result.status == Office.AsyncResultStatus.Succeeded) {
      // Get the File object from the result.
      var myFile = result.value;
      var state = {
        file: myFile,
        counter: 0,
        sliceCount: myFile.sliceCount
      };

      updateStatus("Getting file of " + myFile.size + " bytes");
      getSlice(state);
    } else {
      updateStatus("init " + result.status);
    }
  });
}
function getSlice(state) {
  state.file.getSliceAsync(state.counter, function(result) {
    if (result.status == Office.AsyncResultStatus.Succeeded) {
      updateStatus("Sending piece " + (state.counter + 1) + " of " + state.sliceCount);
      sendSlice(result.value, state);
    } else {
      updateStatus(result);
    }
  });
}

function sendSlice(slice, state) {
  var data = slice.data;

  // If the slice contains data, create an HTTP request.
  if (data) {
    console.log("DATA", data);
    // var fileData = myEncodeBase64(data);

    // var request = new XMLHttpRequest();

    // request.onreadystatechange = function() {
    //   if (request.readyState == 4) {
    //     updateStatus("Sent " + slice.size + " bytes.");
    //     state.counter++;

    //     if (state.counter < state.sliceCount) {
    //       getSlice(state);
    //     } else {
    //       closeFile(state);
    //     }
    //   }
    // };

    // request.open("POST", "");
    // request.setRequestHeader("Slice-Number", slice.index);

    // request.send(fileData);
  }
}
function closeFile(state) {
  state.file.closeAsync(function(result) {
    if (result.status == "succeeded") {
      updateStatus("File closed.");
    } else {
      updateStatus("File couldn't be closed.");
    }
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
      "Response text from server",
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

async function getselectedtext() {
  PowerPoint.run(async function(context) {
    console.log;
    var range = context.presentation.getSelection();
    range = range.getRange("start");
    range.load("font");

    return context.sync().then(function() {
      console.log("Font: " + range.font.name);
      console.log("Size: " + range.font.size);
    });
  });
}
