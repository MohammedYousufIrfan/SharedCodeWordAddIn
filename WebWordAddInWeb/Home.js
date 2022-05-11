﻿//#region Constants

const base64Image = "iVBORw0KGgoAAAANSUhEUgAAAZAAAAEFCAIAAABCdiZrAAAACXBIWXMAAAsSAAALEgHS3X78AAAgAElEQVR42u2dzW9bV3rGn0w5wLBTRpSACAUDmDRowGoj1DdAtBA6suksZmtmV3Qj+i8w3XUB00X3pv8CX68Gswq96aKLhI5bCKiM + gpVphIa1qQBcQbyQB / hTJlpOHUXlyEvD885vLxfvCSfH7KIJVuUrnif + z7nPOd933v37h0IIWQe + BEvASGEgkUIIRQsQggFixBCKFiEEELBIoRQsAghhIJFCCEULEIIBYsQQihYhBBCwSKEULAIIYSCRQghFCxCCAWLEEIoWIQQQsEihCwQCV4CEgDdJvYM9C77f9x8gkyJV4UEznvs6U780rvAfgGdg5EPbr9CyuC1IbSEJGa8KopqBWC / gI7Fa0MoWCROHJZw / lxWdl3isITeBa8QoWCRyOk2JR9sVdF + qvwnnQPsF + SaRSEjFCwSCr0LNCo4rYkfb5s4vj / h33YOcFSWy59VlIsgIRQs4pHTGvYMdJvIjupOx5Ir0Tjtp5K / mTKwXsSLq2hUWG0R93CXkKg9oL0 + ldnFpil + yhlicIM06NA2cXgXySyuV7Fe5CUnFCziyQO2qmg8BIDUDWzVkUiPfHY8xOCGT77EWkH84FEZbx4DwOotbJpI5nj5CQWLTOMBj8votuRqBWDP8KJWABIr2KpLwlmHpeHKff4BsmXxFQmhYBGlBxzoy7YlljxOcfFAMottS6JH + 4Xh69IhEgoWcesBNdVQozLyd7whrdrGbSYdIqFgkQkecMD4epO9QB4I46v4tmbtGeK3QYdIKFhE7gEHjO / odSzsfRzkS1 + 5h42q + MGOhf2CuPlIh0goWPSAogcccP2RJHI1riP + kQYdVK9Fh0goWPSAk82a5xCDG4zPJaWTxnvSIVKwKFj0gEq1go8QgxtUQQeNZtEhUrB4FZbaA9pIN + 98hhhcatbNpqRoGgRKpdAhUrDIMnpAjVrpJSNApK / uRi7pEClYZIk84KDGGQ + IBhhicMP6HRg1ycedgVI6RELBWl4POFCr8VWkszpe3o76G1aFs9ws + dMhUrDIInvAAeMB0ZBCDG6QBh2kgVI6RAoWWRYPqBEI9 + oQEtKgg3sNpUOkYJGF8oADxgOioUauXKIKOkxV99EhUrDIgnhAG + mCUQQhBpeaNb4JgOn3AegQKVhkvj2gjXRLLrIQgxtUQYdpNYsOkYJF5tUDarQg4hCDS1u3VZd83IOw0iFSsMiceUCNWp3WYH0Wx59R6ls9W1c6RAoWmQ8PaCNdz55hiMEN4zsDNhMDpXSIFCwylx5Qo1a9C3yVi69a2ajCWZ43NOkQKVgkph5wwHi + KQ4hBs9SC9 + RMTpEChaJlwfUFylWEafP5uMKqIIOPv0sHSIFi8TFAzpLiXxF / KCbdetEGutFUSa6TXQsdKypv42UgZQhfrWOhbO6q8nPqqCD / zU4OkQKFpm9B7SRbrTpQwzJHNaL / VHyiRVF0dfC2xpOzMnKlUgjW0amhGRW / ZM + w5sqzuqTNWtb9nKBZDLoEClYZGYe0EYaENWHGDaquHJv5CPnz / H9BToWkjmsFkTdOX0GS22p1ovYNEdUr9vCeR3dJlIG1gojn2o8RKPiRX + D0iw6RAoWmYEH1HioiQZqq47VW32dalUlfi1fQf7ByEdUQpMpYfOJ46UPcFweKaMSaWyaWL8z / Mibxzgqe3G4CC6pT4dIwSLReUCNWrkJMdjh8sMSuk1d3bReRGb3hy97iS / SEl + 5bQ0LqM4B9gvytaptC6kbwz++vD3ZG0r3EBDoWUg6RAoWCd0D9isXReTKTYghZbhdUB / UYlKV2TSHitZtYc9QrqynDGy / GnGg + 4XJr779ShJ0gNdAKR3i / PAjXoIZe8BGBS + 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";

// #endregion

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function () {

        // #region Document Ready Region

        $(document).ready(function () {

            // If not using Word 2016, use fallback logic.
            if (Office.context.requirements.isSetSupported('WordApi', '1.3')) {

                // #region API 1.3

                $('#btnAlignTextRight').click(AlignTextRight_New);
                $('#btnApplyInBuildStyle').click(ApplyInBuildStyle_New);
                $("#btnChangeFont").click(ChangeFont_New);
                $("#btnInsertAbbrevation").click(InsertTextIntoRange_New);
                $("#btnAddVersionInfo").click(InsertTextBeforeRange_New);
                $("#btnReplaceText").click(ReplaceText_New);
                $("#btnAddHtml").click(InsertHTML_New);
                $("#btnAddTable").click(InsertTable_New);
                $("#btnCreateContentControl").click(CreateContentControl_New);
                $("#btnReplaceContentControl").click(ReplaceContentInControl_New);
                $("#btnAddBullet").click(Bullet_New);
                $("#OfficeVersion").html("This code is using word 2019 or later");

                // #endregion 
            }

            else if (Office.context.requirements.isSetSupported('WordApi', '1.1')) {

                // #region API 1.1

                $('#btnAlignTextRight').click(AlignTextRight_Old);
                $('#btnApplyInBuildStyle').click(ApplyInBuildStyle_Old);
                $("#btnChangeFont").click(ChangeFont_Old);
                $("#btnInsertAbbrevation").click(InsertTextIntoRange_Old);
                $("#btnAddVersionInfo").click(InsertTextBeforeRange_Old);
                $("#btnReplaceText").click(ReplaceText_Old);
                $("#btnAddHtml").click(InsertHTML_Old);
                $("#btnAddTable").click(InsertTable_Old);
                $("#btnCreateContentControl").click(CreateContentControl_Old);
                $("#btnReplaceContentControl").click(ReplaceContentInControl_Old);
                $("#OfficeVersion").html("This code is using word 2016 or later");
                $("#btnAddBullet").click(Bullet_New);
                // #endregion 

            }

            // #region Common

            $('#btnAddParagraph').click(AddParagraph);
            $('#btnAddImage').click(AddImage);

            // #endregion

            // #region Api Methods

            $("#btnShowUnicode").click(GetUnicode);
            $("#btnShowCharCount").click(GetCharCount);
            $("#btnShowWordCount").click(GetWordCount);

            // #endregion
        });

        // #endregion 
    };

})();


// #region APi Interaction Code


function GetUnicode() {
    Word.run(function (context) {
        const range = context.document.getSelection();
        context.load(range, 'text');
        return context.sync().then(function () {
            const url = "https://localhost:44324/wordanalyzer/unicode?value=" + range.text;
            $.ajax({
                type: "GET",
                url: url,
                success: function (data) {
                    let htmlData = data.replace(/\r\n/g, '<br>');
                    $("#txtUnicodeResult").html(htmlData);
                },
                error: function (data) {
                    $("#txtUnicodeResult").html("error occurred in ajax call.");
                }
            });
        });
    });
}

function GetCharCount() {
    Word.run(function (context) {
        const range = context.document.getSelection();
        context.load(range, 'text');
        return context.sync().then(function () {
            const url = "https://localhost:44324/wordanalyzer/charcount?value=" + range.text;
            $.ajax({
                type: "GET",
                url: url,
                success: function (data) {
                    let htmlData = data.replace(/\r\n/g, '<br>');
                    $("#txtCharCountResult").html(htmlData);
                },
                error: function (data) {
                    $("#txtCharCountResult").html("error occurred in ajax call.");
                }
            });
        });
    });
}

function GetWordCount() {
    Word.run(function (context) {
        const range = context.document.getSelection();
        context.load(range, 'text');
        return context.sync().then(function () {
            const url = "https://localhost:44324/wordanalyzer/wordcount?value=" + range.text;
            $.ajax({
                type: "GET",
                url: url,
                success: function (data) {
                    let htmlData = data.replace(/\r\n/g, '<br>');
                    $("#txtWordCountResult").html(htmlData);
                },
                error: function (data) {
                    $("#txtWordCountResult").html("error occurred in ajax call.");
                }
            });
        });
    });
}
// #endregion APi Interaction


// #region Document Interaction Code

// #region Common Methods

function AddParagraph() {
    Word.run(function (context) {
        context.document.body.insertParagraph(
            "Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
            "Start"
        );
        return context.sync();
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function AddImage() {

    Word.run(function (context) {
        context.document.body.insertInlinePictureFromBase64(base64Image, "End");
        return context.sync();
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

// #endregion

// #region Api 1.3 Methods

function AlignTextRight_New() {
    Word.run(function (context) {
        var firstParagraph = context.document.body.paragraphs.getFirst();
        return context.sync().then(function () {
            firstParagraph.alignment = Word.Alignment.right;
        });
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function ApplyInBuildStyle_New() {
    Word.run(function (context) {
        var firstParagraph = context.document.body.paragraphs.getFirst();
        return context.sync().then(function () {
            firstParagraph.styleBuiltIn = "Emphasis";
        });
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function ChangeFont_New() {

    Word.run(function (context) {
        var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
        return context.sync().then(function () {
            secondParagraph.font.set({
                name: "Courier New",
                bold: true,
                size: 18,
            });
        });
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function InsertTextIntoRange_New() {

    Word.run(function (context) {
        var originalRange = context.document.getSelection();
        originalRange.insertText("(C2R)", "End");
        originalRange.load("text");
        return context.sync().then(function () {
            context.document.body.insertParagraph("Original range: " + originalRange.text, "End");

        });
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function InsertTextBeforeRange_New() {

    Word.run(function (context) {
        var originalRange = context.document.getSelection();
        originalRange.insertText("Office 2016, ", "Before");
        originalRange.load("text");
        return context.sync().then(function () {
            context.document.body.insertParagraph("Original range: " + originalRange.text, "End");
        });
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function ReplaceText_New() {

    Word.run(function (context) {
        var originalRange = context.document.getSelection();
        originalRange.insertText("many", "Replace");
        return context.sync();
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function InsertHTML_New() {

    Word.run(function (context) {
        var blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
        blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
        return context.sync();
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function InsertTable_New() {

    Word.run(function (context) {
        const tableData = [
            ["Name", "ID", "Birth City"],
            ["Bob", "434", "Chicago"],
            ["Sue", "719", "Havana"],
        ];
        var blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
        blankParagraph.insertTable(3, 3, "After", tableData);

        return context.sync();
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}
function Bullet_New1() {

    Word.run(function (context) {

        let paragraphs = context.document.body.paragraphs;
        paragraphs.load("items");

        var secondParagraph;
        return context.sync().then(function () {
            if (paragraphs.items.length > 1) {
                secondParagraph = paragraphs.items[1];
                secondParagraph.load("font");
            }
            let list = paragraphs.items[1].startNewList();

            // insert the list at the start location
             list.insertParagraph("My 0 item at start1", Word.InsertLocation.start)
        }).then(context.sync).then(function () {
            secondParagraph.font.set({
                name: "Courier New",
                bold: true,
                size: 30,
            });
             list.insertParagraph("23", Word.InsertLocation.start)
        })
        //return context.sync();
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function Bullet_New() {

    Word.run(function (context) {
        context.document.body.insertParagraph(
            "Mohammed1 Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
            "Start"
        );
        let paragraphs = context.document.body.paragraphs;
        paragraphs.load("items");
        return context.sync().then(function () {
            let list = paragraphs.items[1].startNewList();
            // insert the list at the start location
            list.insertParagraph("My 0 item at start2", Word.InsertLocation.start);

        }).then(context.sync).then(function () {
            let wordLists = context.document.body.lists;

            // grab the first list
            let firstList = wordLists.getFirstOrNullObject();

            // set the bullet design for the first level
            firstList.setLevelBullet(0, Word.ListBullet.checkmark);

            // set indent level for the first level to 50 points, 20 points for images
            firstList.setLevelIndents(0, 50, 20);

        })
        //return context.sync();
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}
function CreateContentControl_New() {

    Word.run(function (context) {
        let serviceNameRange = context.document.getSelection();
        let serviceNameContentControl = serviceNameRange.insertContentControl();
        serviceNameContentControl.title = "Service Name";
        serviceNameContentControl.tag = "serviceName";
        serviceNameContentControl.appearance = "Tags";
        serviceNameContentControl.color = "blue";
        return context.sync();
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function ReplaceContentInControl_New() {

    Word.run(function (context) {
        var serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
        serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");

        return context.sync();

    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}
// #endregion 

// #region Api 1.1 Methods

function AlignTextRight_Old() {

    Word.run(function (context) {
        let paragraphs = context.document.body.paragraphs;
        paragraphs.load("text");
        var firstParagraph;
        return context.sync().then(function () {
            if (paragraphs.items.length > 0) {
                firstParagraph = paragraphs.items[0];
                firstParagraph.load("alignment");
            }
        }).then(context.sync).then(function () {
            firstParagraph.alignment = Word.Alignment.right;
        })
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function ApplyInBuildStyle_Old() {

    Word.run(function (context) {
        let paragraphs = context.document.body.paragraphs;
        paragraphs.load("text");
        paragraphs.styleBuiltIn = Word.Style.intenseReference;
        var firstParagraph;
        return context.sync().then(function () {
            if (paragraphs.items.length > 0) {
                firstParagraph = paragraphs.items[0];
                firstParagraph.load("styles");
            }
        }).then(context.sync).then(function () {
            firstParagraph.style = "Emphasis";
        })
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function ChangeFont_Old() {

    Word.run(function (context) {
        let paragraphs = context.document.body.paragraphs;
        paragraphs.load("text");
        var secondParagraph;
        return context.sync().then(function () {
            if (paragraphs.items.length > 1) {
                secondParagraph = paragraphs.items[1];
                secondParagraph.load("font");
            }
        }).then(context.sync).then(function () {
            secondParagraph.font.set({
                name: "Courier New",
                bold: true,
                size: 18,
            });
        })
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}


function InsertTextIntoRange_Old() {

    Word.run(function (context) {
        var originalRange = context.document.getSelection();
        context.load(originalRange, 'text');
        return context.sync().then(function () {
            originalRange.insertText(" (C2R)", "End");
            originalRange.load("text");

        }).then(context.sync).then(function () {
            context.document.body.insertParagraph("Original range: " + originalRange.text, "End");
        })
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function InsertTextBeforeRange_Old() {

    Word.run(function (context) {
        var originalRange = context.document.getSelection();
        context.load(originalRange, 'text');
        return context.sync().then(function () {
            originalRange.insertText("Office 2016, ", "Before");
            originalRange.load("text");

        }).then(context.sync).then(function () {
            context.document.body.insertParagraph("Original range: " + originalRange.text, "End");
        })
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function ReplaceText_Old() {

    Word.run(function (context) {
        var originalRange = context.document.getSelection();
        context.load(originalRange, 'text');
        return context.sync().then(function () {
            originalRange.insertText("many", "Replace");
        });
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function InsertHTML_Old() {

    Word.run(function (context) {
        let paragraphs = context.document.body.paragraphs;
        paragraphs.load("text");
        var blankParagraph;
        return context.sync().then(function () {
            if (paragraphs.items.length > 0) {
                blankParagraph = paragraphs.items[paragraphs.items.length - 1].insertParagraph("", "After");
                blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
            }
        });
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function InsertTable_Old() {

    Word.run(function (context) {
        let paragraphs = context.document.body.paragraphs;
        paragraphs.load("text");
        var blankParagraph;
        const tableData = [
            ["Name", "ID", "Birth City"],
            ["Bob", "434", "Chicago"],
            ["Sue", "719", "Havana"],
        ];
        return context.sync().then(function () {
            if (paragraphs.items.length > 0) {
                //blankParagraph = paragraphs.items[paragraphs.items.length - 1].insertParagraph("", "After");
                blankParagraph = paragraphs.items[paragraphs.items.length - 1];
                blankParagraph.insertTable(3, 3, "After", tableData);
            }
        })
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function CreateContentControl_Old() {

    Word.run(function (context) {
        let serviceNameRange = context.document.getSelection();
        context.load(serviceNameRange, 'text');
        return context.sync().then(function () {
            let serviceNameContentControl = serviceNameRange.insertContentControl();
            serviceNameContentControl.title = "Service Name";
            serviceNameContentControl.tag = "serviceName";
            serviceNameContentControl.appearance = "Tags";
            serviceNameContentControl.color = "blue";
        });
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

function ReplaceContentInControl_Old() {

    Word.run(function (context) {
        let doc = context.document;
        doc.load("contentControls");
        return context.sync().then(function () {
            var serviceNameContentControl = doc.contentControls.getByTag("serviceName").items[0];
            serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
        });
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

// #endregion


// #endregion