//Authoer Lai Min Han

function startMenu() {
  // Get the user interface of the spreadsheet
  const ui = SpreadsheetApp.getUi();
  // Create a new menu call "Spam Solution Challenge"
  ui.createMenu('Solution Challenge')
    //Add an item to the menu that will call the "Spam Solution Challenge" function when clicked
    .addItem('First Spam', 'spam')
    .addToUi()
}

// use this to convert word to html "https://wordtohtml.net/"
function spam() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getDataRange().getValues(); // Get all data in the sheet
  const current = data[1][8]; //this will be the input column for them to detect which one to blast
  // Start from row 2 (assuming row 1 is headers)
    for (var i = 1; i < data.length; i++) {
      var cellToUpdate = sheet.getRange(i + 1, 8); // Column 8 (index 7) for chip, row i+1
      if (data[i][7]===current){ //to check same with your input to be sent
        var email = data[i][2]; // Assuming emails are in the 3th index (column C)
        if (email) { // Check if the email address is valid
          var subject = "Ready to Solve Real-World Challenges with Google Technology?"; // Set email subject
          var htmlBody = `
          <p><span style="color:#222222;font-size:11pt;font-family:'Google Sans',sans-serif;"><span style="border:none;"><img src="https://lh7-rt.googleusercontent.com/docsz/AD_4nXfAvSeDM9TjF0L3On3R0eRqHDpU60Le9JhCSBBoH8N7E4QtSY6dxbBU_ZhC9VhUF4GrfR5BqsNCZoRStZOQdU5EU3sq7g4Tc54qMmwxb0r3Q4tT-FrIBO3ZHNTpb9Vo43mBVC9M?key=7kpiPacG1qO66JZpLAnyN1Kc" width="581" height="190.94322916666667"></span></span></p>
    <p><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;">Hi there,</span><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;"><br><br></span></p>
    <p><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;">Do you see challenges in your community that technology could solve?</span></p>
    <p><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;"><br></span><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;">GDG on Campus APAS team invites you to participate in APAC Solution Challenge 2025, powered by Hack2skill to give you an opportunity to solve local problems by building solutions using Google AI technologies.&nbsp;</span></p>
    <p><br></p>
    <p><strong><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;">What is APAC Solution Challenge 2025?</span></strong></p>
    <p><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;">The&nbsp;</span><a href="https://vision.hack2skill.com/event/apacsolutionchallenge?utm_source=hack2skill&utm_medium=homepage"><strong><u><span style="color:#1155cc;font-size:10pt;font-family:'Google Sans',sans-serif;">APAC Solution Challenge 2025</span></u></strong></a><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;">&nbsp;is an opportunity for&nbsp;</span><strong><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;">student developers</span></strong><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;">&nbsp;to tackle pressing issues&mdash;such as agriculture, tourism, trade, healthcare and sustainability&mdash;using one or more Google AI technologies. By participating, student developers not only gain invaluable skills but also contribute to making a positive impact on society using technology.&nbsp;</span></p>
    <div align="center">
        <table style="border: none; border-collapse: collapse; width: 100%;">
            <tbody>
                <tr>
                    <td style="color:#00a848;background-color:#00a848;">
                        <p style="text-align: center;"><a href="https://vision.hack2skill.com/event/apacsolutionchallenge?source=hack2skill&medium=homepage&utm_source=hack2skill&utm_medium=homepage"><strong><u><span style="color:#ffffff;font-size:15pt;font-family:'Google Sans',sans-serif;">Click here to Register Now</span></u></strong></a></p>
                    </td>
                </tr>
            </tbody>
        </table>
    </div>
    <p><br></p>
    <p><strong><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;">What kind of problems can you solve for the Hackathon?</span></strong><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;"><br></span><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;">You can tackle local challenges in areas such as agriculture, tourism, trade, healthcare, and sustainability&mdash;leveraging&nbsp;</span><strong><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;">one or more Google AI technologies&nbsp;</span></strong><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;">to create impactful solutions.</span></p>
    <div align="center">
        <table style="border: none; border-collapse: collapse; width: 100%;">
            <tbody>
                <tr>
                    <td style="color:#1155cc;background-color:#1155cc;border: solid #d9d9d9 1pt;">
                        <p style="text-align: center;"><a href="https://vision.hack2skill.com/event/apacsolutionchallenge?source=hack2skill&medium=homepage&utm_source=hack2skill&utm_medium=homepage"><strong><u><span style="color:#ffffff;font-size:12pt;font-family:'Google Sans',sans-serif;">Click here to explore Hackathon Themes</span></u></strong></a></p>
                    </td>
                </tr>
            </tbody>
        </table>
    </div>
    <p><br></p>
    <p><strong><span style="color:#434343;font-size:12pt;font-family:'Google Sans',sans-serif;">Timeline for APAC Solution Challenge:</span></strong><span style="color:#222222;background-color:#ffffff;font-size:11pt;font-family:Arial,sans-serif;"><span style="border:none;"><img alt="⏱️" src="https://lh7-rt.googleusercontent.com/docsz/AD_4nXd88j57ukL-7--F2keAKajG_YxaFRlaqG-ThmCCU2edzg0Et3LdGdXgnymh-vemdRkjt5Iz7v7Tp3r01Qsgj14vq4Ilf4EbHcBVz_dOnXXLAYox0nhO7tgiE13natAREyyxo7bb?key=7kpiPacG1qO66JZpLAnyN1Kc" width="17" height="17"></span></span><span style="font-size:11pt;font-family:'Google Sans',sans-serif;"><br></span><span style="font-size:11pt;font-family:'Google Sans',sans-serif;"><span style="border:none;"><img src="https://lh7-rt.googleusercontent.com/docsz/AD_4nXe-m8yYnXu4rK7-q3MWDa3rbERmMYvkHllYnrsjrykYlry1yD0ekbEtMMKuor_fCsghoQF58kFw-MqRCnpEpiVMhyaRzQTTwC_7nlelbtpUMRaAQOgLwMZlqLc9VojQOBjVHLa3?key=7kpiPacG1qO66JZpLAnyN1Kc" width="581" height="256"></span></span></p>
    <p><em><span style="color:#434343;font-size:10pt;font-family:'Google Sans',sans-serif;">Note : Timelines are subject to change**</span></em></p>
    <p><strong><span style="color:#434343;font-size:12pt;font-family:'Google Sans',sans-serif;">Why Participate?</span></strong><span style="color:#222222;background-color:#ffffff;font-size:11pt;font-family:Arial,sans-serif;"><span style="border:none;"><img alt="⚡" src="https://lh7-rt.googleusercontent.com/docsz/AD_4nXfRHQCtTsLWsYybeSzdpLPP8g8RPxYu39BtD4-gQ-l-f5ecN8LVReiIpoRPw9C88aZ8nWFloVP2NBX-uzfbzj5CW7t9YD97nFK9WqICZzojBJJOGtaWPkkCiumAhb0e2A4Y2iv7?key=7kpiPacG1qO66JZpLAnyN1Kc" width="17" height="17"></span></span></p>
    <div align="center">
        <table style="border: none; border-collapse: collapse; width: 100%;">
            <tbody>
                <tr>
                    <td style="border: solid #666666 1pt;">
                        <p style="text-align: center;"><span style="color:#434343;font-size:13pt;font-family:'Google Sans',sans-serif;"><span style="border:none;"><img src="https://lh7-rt.googleusercontent.com/docsz/AD_4nXd7MV2tu-oBh54R2_I4YjTHVyXGdxyyVopc4jqUA-3_cNTpF5V0icouE-X1O5TsKDd108xCQsbAnnIWSGEGx45xS2BDLCM0on79biTvDBLrTxNOu-Jm11AOm0xHEOKmfKVWwURX?key=7kpiPacG1qO66JZpLAnyN1Kc" width="55" height="55"></span></span></p>
                    </td>
                    <td style="border: solid #666666 1pt;">
                        <p style="text-align: center;"><strong><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;">Tackle real-world problems</span></strong><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;">&nbsp;using Google technology</span></p>
                    </td>
                    <td style="border: solid #666666 1pt;">
                        <p style="text-align: center;"><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;"><span style="border:none;"><img src="https://lh7-rt.googleusercontent.com/docsz/AD_4nXf48JTpQ0q6RWUbhts39CPe-9nWi9cQ0PAh2yWnSEmhLOlWoBjE9Eoiv8fCpciOSwkTyCLDl1Hq_QZEA6qoidgqd05Ms7lWgb9btLaUat-UrMESNkjEfZWH4Y9KkZq1zyK1NZs6?key=7kpiPacG1qO66JZpLAnyN1Kc" width="64" height="64"></span></span></p>
                    </td>
                    <td style="border: solid #666666 1pt;">
                        <p style="text-align: center;"><span style="color:#444746;background-color:#ffffff;font-size:11pt;font-family:'Google Sans',sans-serif;">Opportunity to&nbsp;</span><strong><span style="color:#444746;background-color:#ffffff;font-size:11pt;font-family:'Google Sans',sans-serif;">network with student developers</span></strong><span style="color:#444746;background-color:#ffffff;font-size:11pt;font-family:'Google Sans',sans-serif;">&nbsp;across the country</span></p>
                    </td>
                </tr>
                <tr>
                    <td style="border: solid #666666 1pt;">
                        <p style="text-align: center;"><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;"><span style="border:none;"><img src="https://lh7-rt.googleusercontent.com/docsz/AD_4nXeY65qEpxvrbP8IatrWsDxPMvJ-zf7o-gd26r_yV-fdadAI38KK4T2xJ9xWe53AjW7ZpU9poJ1twyoaiQYi5Hy1CjOGzLao_GbXWEMN0Yv2ZGtQRKX5vYXnPS2KSqwsVJQL8FD4?key=7kpiPacG1qO66JZpLAnyN1Kc" width="59" height="59"></span></span></p>
                    </td>
                    <td style="border: solid #666666 1pt;">
                        <p style="text-align: center;"><strong><span style="color:#444746;font-size:11pt;font-family:'Google Sans',sans-serif;">&nbsp;Online training sessions</span></strong></p>
                    </td>
                    <td style="border: solid #666666 1pt;">
                        <p style="text-align: center;"><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;"><span style="border:none;"><img src="https://lh7-rt.googleusercontent.com/docsz/AD_4nXe3uxrLelZ5ICf56--gqiWlC6e_1Q0-U26NYC6BH9UlfOtllq8Y05QyGO6ONZ__YAWYmNAcr_n9tWvAWAY2S-rvoG2gsKQb9nTOpB0PObJFpr7y7nb0qWdfjnuMV06SqMfIQqAz?key=7kpiPacG1qO66JZpLAnyN1Kc" width="70" height="60"></span></span></p>
                    </td>
                    <td style="border: solid #666666 1pt;">
                        <p style="text-align: center;"><strong><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;">Recognition and rewards&nbsp;</span></strong><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;">with a cash prize pool of&nbsp;</span><strong><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;">5000 USD</span></strong></p>
                    </td>
                </tr>
            </tbody>
        </table>
    </div>
    <p><strong><span style="color:#434343;font-size:12pt;font-family:'Google Sans',sans-serif;">&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</span></strong></p>
    <div align="center">
        <table style="border: none; border-collapse: collapse; width: 100%;">
            <tbody>
                <tr>
                    <td style="color:#6aa84f;background-color:#6aa84f;border: solid #999999 1pt;">
                        <p style="text-align: center;"><a href="https://vision.hack2skill.com/event/apacsolutionchallenge?source=hack2skill&medium=homepage&utm_source=hack2skill&utm_medium=homepage"><strong><u><span style="color:#ffffff;font-size:12pt;font-family:'Google Sans',sans-serif;">Ready to create an impact with AI by solving local challenges?</span></u></strong></a></p>
                        <p style="text-align: center;"><a href="https://vision.hack2skill.com/event/apacsolutionchallenge?source=hack2skill&medium=homepage&utm_source=hack2skill&utm_medium=homepage"><strong><u><span style="color:#ffffff;font-size:12pt;font-family:'Google Sans',sans-serif;">Join the APAC Solution Challenge 2025!</span></u></strong></a></p>
                    </td>
                </tr>
            </tbody>
        </table>
    </div>
    <p><br></p>
    <p><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;">In case of any queries, read the&nbsp;</span><a href="https://docs.google.com/document/d/17D-ukO4000yODLU2eO27DOyrEfgpEmEwewRnzAU2uRQ/edit?usp=drive_link"><strong><u><span style="color:#1155cc;font-size:11pt;font-family:'Google Sans',sans-serif;">FAQs here</span></u></strong></a><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;">&nbsp;or join the&nbsp;</span><a href="https://discord.gg/sYY9epGdKY"><strong><u><span style="color:#1155cc;font-size:11pt;font-family:'Google Sans',sans-serif;">Discord server</span></u></strong></a><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;">&nbsp;or mail us at&nbsp;</span><a href="mailto:apacsolutionchallenge@hack2skill.com"><u><span style="color:#1155cc;font-size:11pt;font-family:'Google Sans',sans-serif;">apacsolutionchallenge@hack2skill.com</span></u></a><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;">.</span></p>
    <p style="text-align: center;"><a href="https://twitter.com/hack2skill"><strong><u><span style="color:#1155cc;font-size:11pt;font-family:'Google Sans',sans-serif;"><span style="border:none;"><img src="https://lh7-rt.googleusercontent.com/docsz/AD_4nXcH4kp4i2OAV7gHa-wbA6MysESqW9EtQw8xiS99YDGI30j3ohAiIT482QA1b36_ASMZg4GID0U7-fsdg8w8Noze1IrKH0_q_5nHN6MzwYJwjroJG4Xdn6F7If4jFC6DcAH4fbI?key=7kpiPacG1qO66JZpLAnyN1Kc" width="34" height="22"></span></span></u></strong></a><strong><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;">|&nbsp;</span></strong><a href="https://www.linkedin.com/company/hack2skill/"><u><span style="color:#1155cc;font-size:11pt;font-family:'Google Sans',sans-serif;"><span style="border:none;"><img src="https://lh7-rt.googleusercontent.com/docsz/AD_4nXeiCi-PM8w5oZ9yMHaB1Oo_hSFsbJDdwFV546t29tgS7p6Xju-xSklHqrzg_lJ_t4BCKaLrn1dOnlydlmAFZFkZHcdbh3_MOPL3SU3_WvRZ7OvdT9O2qpNVoaERCxiWTIHYCqFz?key=7kpiPacG1qO66JZpLAnyN1Kc" width="23" height="23"></span></span></u></a><span style="color:#434343;font-size:11pt;font-family:'Google Sans',sans-serif;">|&nbsp;</span><a href="https://www.instagram.com/hack2skill/"><u><span style="color:#1155cc;font-size:11pt;font-family:'Google Sans',sans-serif;"><span style="border:none;"><img src="https://lh7-rt.googleusercontent.com/docsz/AD_4nXenw8N3moqAWstiQ3L3R61FjcRqVp_6WEIlUrfi_1SEYhuMtQ83ivGJeZnI4u2VtByPgxc66Ay9hiG5VyTlj44zp_6szcqQfkx5vtEiqo0JFbPKV_VpGXP09mhMCTY2CL_M2m-m?key=7kpiPacG1qO66JZpLAnyN1Kc" width="27" height="27.09190004559216"></span></span></u></a></p>
          `;
          try {
            MailApp.sendEmail({
              to: email,
              subject: subject,
              htmlBody: htmlBody,
            });
            Logger.log("Email sent to " + email);
            // Update cell to show how much time this had sent
            cellToUpdate.setValue((current+1));
            //data[i][7]=current+1;
            Utilities.sleep(1000); // 1 second delay

          } catch (e) {
            Logger.log("Error sending email to " + email + ": " + e.toString());
          }
        }
      }
    }
}

