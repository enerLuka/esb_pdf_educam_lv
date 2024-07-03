<script>
    import { read, utils } from "xlsx";
    import { PDFDocument, rgb, StandardFonts } from "pdf-lib";
    import Dropzone from "svelte-file-dropzone";
    import JSZip from "jszip";

    let excelData = [];
    let selectedEmployees = [];
    let pdfFile = null;
    let modifiedPdfs = [];
    let excelFiles = {
        accepted: [],
        rejected: [],
    };
    let pdfFiles = {
        accepted: [],
        rejected: [],
    };

    $: console.log(selectedEmployees);
    // Functie om het Excel-bestand te lezen
    async function handleExcel(file) {
        console.log("file:", file);
        const data = await file.arrayBuffer();
        const workbook = read(data, { type: "array" });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        // Convert the worksheet to a JSON array (including header)
        let sheetData = utils.sheet_to_json(firstSheet, {
            header: 1,
            raw: false,
        });

        // Get the headers from the second row
        const headers = sheetData[1];

        // Create the array of objects using the headers from the second row, and skipping the first row
        const dataWithoutFirstRow = sheetData.slice(2).map((row) => {
            let obj = {};
            headers.forEach((header, index) => {
                obj[header] = row[index];
            });
            return obj;
        });

        console.log(dataWithoutFirstRow);
        excelData = dataWithoutFirstRow;
    }

    async function handleExcelFileSelect(e) {
        const { acceptedFiles, fileRejections } = e.detail;
        excelFiles.accepted = [...excelFiles.accepted, ...acceptedFiles];
        excelFiles.rejected = [...excelFiles.rejected, ...fileRejections];
        console.log(excelFiles);
        await handleExcel(excelFiles.accepted[0]);
    }

    async function handlePdfFileSelect(e) {
        const { acceptedFiles, fileRejections } = e.detail;
        pdfFiles.accepted = [...pdfFiles.accepted, ...acceptedFiles];
        pdfFiles.rejected = [...pdfFiles.rejected, ...fileRejections];
        pdfFile = await pdfFiles.accepted[0].arrayBuffer();
    }

    function splitStringAtNumber(input) {
        if (!input) {
            return "";
        }
        // Regular expression to find the first occurrence of a number
        const regex = /\d/;
        const match = input.match(regex);

        if (match) {
            const index = match.index;
            const firstPart = input.substring(0, index).trim();
            const secondPart = input.substring(index).trim();
            return [firstPart, secondPart];
        } else {
            // If no number is found, return the original string in an array
            return [input, ""];
        }
    }

    function convertDateString(dateString) {
        if (!dateString) {
            return "";
        }
        // Split the date string by the '/' character
        const [month, day, year] = dateString.split("/");

        // Pad single-digit day and month with leading zeros
        const paddedDay = day.padStart(2, "0");
        const paddedMonth = month.padStart(2, "0");
        const paddedYear = year.padStart(2, "0");

        // Concatenate the parts into the desired format
        return `${paddedDay}${paddedMonth}${paddedYear}`;
    }

    function splitSocialString(str) {
        if (!str) {
            return "";
        }
        // Remove non-alphanumeric characters using regex
        const cleanStr = str.replace(/\W/g, "");

        // Extract parts based on specified lengths
        const part1 = cleanStr.substring(0, 6);
        const part2 = cleanStr.substring(6, 9);
        const part3 = cleanStr.substring(9, 11);

        return [part1, part2, part3];
    }

    function formatPhoneNumber(phoneNumber) {
        if (!phoneNumber) {
            return "";
        }
        // Remove all non-numeric characters using regex
        const cleaned = phoneNumber.replace(/\D/g, "");

        // Check if the cleaned number starts with '32' (Belgium country code) and remove it if present
        let formatted = cleaned.startsWith("32")
            ? cleaned.substring(2)
            : cleaned;

        // If the number starts with '0', remove the leading '0'
        if (formatted.startsWith("0")) {
            formatted = formatted.substring(1);
        }

        return formatted;
    }

    // Functie om geselecteerde rijen in de PDF's te vullen
    async function fillPDFs() {
        if (!pdfFile || selectedEmployees.length === 0) return;

        modifiedPdfs = [];
        for (const employee of selectedEmployees) {
            const pdfDoc = await PDFDocument.load(pdfFile);
            const pages = pdfDoc.getPages();
            const firstPage = pages[0];

            const employeeData = excelData.find(
                (item) => item["GUID"] == employee,
            );

            // Functie om undefined te checken en om te zetten naar een lege string
            const safeText = (text) => (text !== undefined ? text : "");

            // Functie om hoofdletters te gebruiken
            const toUpperCase = (text) =>
                typeof text === "string" ? text.toUpperCase() : text;

            // // Functie om meer letter-spacing toe te voegen
            // const letterSpacing = (text) => text.split("").join("  ");

            const courierFont = await pdfDoc.embedFont(StandardFonts.Courier);

            // Functie om tekst met letter-spacing toe te voegen
            const drawText = (
                page,
                text,
                x,
                y,
                spacing = 1.1,
                size = 12,
                color = rgb(0, 0, 0),
                font = courierFont,
            ) => {
                const letters = toUpperCase(safeText(text)).split("");
                let currentX = x;
                for (const letter of letters) {
                    page.drawText(letter, {
                        x: currentX,
                        y: y, // Y-waarde met 15 verlagen
                        size,
                        color,
                    });
                    currentX += size + spacing; // Vergroot x-coördinaat voor letter-spacing
                }
            };

            drawText(
                firstPage,
                employeeData.LastName,
                70,
                firstPage.getHeight() - 243,
            );
            drawText(
                firstPage,
                employeeData.FirstName,
                88,
                firstPage.getHeight() - 265,
            );
            drawText(
                firstPage,
                formatPhoneNumber(employeeData.Mobile1),
                408,
                firstPage.getHeight() - 265,
            );

            const streetAndNumber = splitStringAtNumber(employeeData.Street1);
            drawText(
                firstPage,
                streetAndNumber[0],
                70,
                firstPage.getHeight() - 286,
            );
            drawText(
                firstPage,
                streetAndNumber[1],
                390,
                firstPage.getHeight() - 286,
            );
            drawText(
                firstPage,
                employeeData.Street2,
                483,
                firstPage.getHeight() - 286,
            );

            drawText(
                firstPage,
                employeeData.PostCode,
                83,
                firstPage.getHeight() - 308,
            );
            drawText(
                firstPage,
                employeeData.CityName,
                222,
                firstPage.getHeight() - 308,
            );
            drawText(
                firstPage,
                employeeData.CountryName,
                470,
                firstPage.getHeight() - 308,
            );

            drawText(
                firstPage,
                employeeData.BirthPlace,
                110,
                firstPage.getHeight() - 330,
            );

            drawText(
                firstPage,
                convertDateString(employeeData.BirthDate),
                440,
                firstPage.getHeight() - 330,
            );
            drawText(
                firstPage,
                employeeData.Email1,
                70,
                firstPage.getHeight() - 353,
            );

            const socialNumbers = splitSocialString(
                employeeData.SocialSecurityNumber,
            );
            drawText(
                firstPage,
                socialNumbers[0],
                133,
                firstPage.getHeight() - 374,
            );
            drawText(
                firstPage,
                socialNumbers[1],
                223,
                firstPage.getHeight() - 374,
            );
            drawText(
                firstPage,
                socialNumbers[2],
                276,
                firstPage.getHeight() - 374,
            );

            let bName = "Enercon Services Belgium";
            let bVat = "0806283202";
            let bStreet = "Vlamingveld";
            let bNumber = "43";
            let bPostCode = "8490";
            let bPlace = "Jabbeke";
            let bCountry = "BE";
            let bMail = "training-service-be@enercon.de";
            let bTel = "050350150";

            drawText(firstPage, bName, 70, firstPage.getHeight() - 580, 0.3);
            drawText(firstPage, bVat, 369, firstPage.getHeight() - 580);
            drawText(firstPage, bStreet, 70, firstPage.getHeight() - 598);
            drawText(firstPage, bNumber, 394, firstPage.getHeight() - 598);
            drawText(firstPage, bPostCode, 85, firstPage.getHeight() - 620);
            drawText(firstPage, bPlace, 221, firstPage.getHeight() - 620);
            drawText(firstPage, bCountry, 473, firstPage.getHeight() - 620);
            drawText(firstPage, bMail, 71, firstPage.getHeight() - 643);
            drawText(firstPage, bTel, 369, firstPage.getHeight() - 655);

            const pdfBytes = await pdfDoc.save();
            modifiedPdfs.push({
                employeeData,
                data: new Blob([pdfBytes], { type: "application/pdf" }),
            });
        }

        console.log("we ended here", modifiedPdfs);

        if (modifiedPdfs.length === 0) return;

        const zip = new JSZip();
        modifiedPdfs.forEach(({ employeeData, data }) => {
            zip.file(
                `Educam_invuldocument_${employeeData.FirstName + employeeData.LastName}.pdf`,
                data,
            );
        });

        const zipBlob = await zip.generateAsync({ type: "blob" });
        const zipUrl = URL.createObjectURL(zipBlob);
        const a = document.createElement("a");
        a.href = zipUrl;
        a.download = "modified_pdfs.zip";
        a.click();
        URL.revokeObjectURL(zipUrl);
    }
</script>

<div class="page">
    <div class="container">
        <h1>Educam Formulier Generator</h1>

        <div class="dropzone">
            <Dropzone on:drop={handleExcelFileSelect}>
                <p>Drop hier je Vario Excellijst met werknemers</p>
            </Dropzone>
        </div>
        <ol>
            {#each excelFiles.accepted as item}
                <li>{item.name}</li>
            {/each}
        </ol>
        {#if excelData.length > 0 && pdfFile != null}
            <h2>Selecteer werknemer (ctrl + klik voor meerdere)</h2>
            <select multiple bind:value={selectedEmployees} size="10">
                {#each excelData as employee}
                    <option value={employee["GUID"]}
                        >{employee.LastName + " " + employee.FirstName}</option
                    >
                {/each}
            </select>
            <button on:click={fillPDFs}>Fill PDFs</button>
        {/if}

        <div class="dropzone">
            <Dropzone on:drop={handlePdfFileSelect}>
                <p>Drop hier je Educam PDF formulier</p>
            </Dropzone>
        </div>
        <ol>
            {#each pdfFiles.accepted as item}
                <li>{item.name}</li>
            {/each}
        </ol>
    </div>
</div>

<style>
    /* Algemene stijlen voor de pagina */
    .page {
        font-family: "Roboto", sans-serif;
        background-color: #f4f4f9;
        color: #333;
        margin: 0;
        padding: 0;
        display: flex;
        justify-content: center;
        align-items: center;
        min-height: 100vh;
    }

    h1 {
        font-size: 2rem;
        color: #4a4a4a;
        text-align: center;
    }

    h2 {
        font-size: 1.5rem;
        color: #4a4a4a;
        margin-top: 1.5rem;
        text-align: center;
    }

    .container {
        width: 100%;
        max-width: 800px;
        padding: 20px;
        margin: 0 auto;
        background: #fff;
        border-radius: 10px;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
    }

    .dropzone {
        border: 2px dashed #ccc;
        padding: 40px;
        margin: 20px 0;
        text-align: center;
        background-color: #fafafa;
        border-radius: 10px;
        transition: border-color 0.3s;
    }

    .dropzone:hover {
        border-color: #aaa;
    }

    .dropzone p {
        font-size: 1rem;
        color: #888;
    }

    select {
        width: 100%;
        padding: 10px;
        margin: 10px 0;
        font-size: 1rem;
        border: 1px solid #ccc;
        border-radius: 5px;
        box-sizing: border-box;
    }

    button {
        display: inline-block;
        padding: 10px 20px;
        margin: 10px 5px;
        font-size: 1rem;
        color: #fff;
        background-color: #007bff;
        border: none;
        border-radius: 5px;
        cursor: pointer;
        transition: background-color 0.3s;
    }

    button:hover {
        background-color: #0056b3;
    }

    ul {
        list-style-type: none;
        padding: 0;
    }

    ul li {
        margin: 10px 0;
    }

    ul li a {
        text-decoration: none;
        color: #007bff;
        transition: color 0.3s;
    }

    ul li a:hover {
        color: #0056b3;
    }
</style>
