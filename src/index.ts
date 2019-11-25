#!/usr/bin/env node

const program = require('commander');
import { randomFillSync } from 'crypto';
import { green, red, yellow, cyan, magenta, white } from 'chalk';
import * as http from 'http';
import * as https from 'https';
import * as filesystem from 'fs';

type InputURLData = {inputIndex: number, inputURL: string}
type RequestData = {guid: string, url: string}
type ResponseData = RequestData & {statusCode?: number, location?: string, error?: Error}
type ProcessedURLData = InputURLData & RequestData & {responses: ResponseData[], targetURL: string | undefined, targetRedirectStatusCode: number | undefined}

const CONCURRENT_REQUESTS_DEFAULT = 5;
const REPORT_FILENAME_SUFFIX_DEFAULT = 'redirects';
const USER_AGENT_DEFAULT = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.70 Safari/537.36';
const URL_PROTOCOL_REGEX = /^(?:f|ht)tps?\:\/\//;
const URL_HTTPS_PROTOCOL_REGEX = /^https:\/\//;
const URL_VALIDATION_REGEX = /^((?:f|ht)tps?:\/\/)?[\w.-]+(?:\.[\w\.-]+)+[\w\-\._~:/?#[\]@!\$&'\(\)\*\+,;=.]+$/;

const report: (ProcessedURLData & {
    numberOfRedirects?: number;
    finalStatusCode?: number;
    finalRedirectStatusCode?: number | string;
    finalURL?: string;
    targetURLMatched?: boolean;
    targetStatusMatched?: boolean;
})[] = [];

let didError = false;

program
    .option('-s, --sites <urls>', 'Comma-delimited list of URLs to check redirects', (string: string) => string.split(','), [])
    .option('-t, --targets <urls>', 'Comma-delimited list of target URLs to compare to final redirect URLs', (string: string) => string.split(','), [])
    .option('-c, --codes <codes>', 'Comma-delimited list of target status codes to compare to final redirect status codes', (string: string) => string.split(','), [])
    .option('-p, --prefix <prefix>', 'Prefix to be applied to all sites without a protocol')
    .option('-r, --protocol <protocol>', 'Protocol to be applied to all sites without a protocol')
    .option('-a, --auth <username:password>', 'Provide a username and password for authentication')
    .option('-n, --concurrent <number>', 'Number of request to make concurrently. Defaults to 5')
    .option('-j, --json', 'Generates JSON reports')
    .option('-x, --xlsx', 'Generates XLSX reports')
    .option('-f, --filename <filename>', 'Set the name of the generated report file')
    .parse(process.argv);

const log = {
    missingInputURLs: () => {
        console.error(red(`\n    ERROR: No site URLs were given.\n    Please make sure to include URLs with -s or --sites: ${magenta('redirect-tester -s google.com,facebook.com')}\n`));
    },
    inputAndTargetURLsLengthMismatched: () => {
        console.error(red(`\n    ERROR: The number of input URLs and target URLs (sites and targets) did not match.\n    If you use targets, the list of targets must be the same length as the list of sites.\n    To skip a target for a specific site, just leave an empty spot in the target list.\n    For example, to skip providing a target for facebook.com:\n    ${magenta('redirect-tester -s google.com,facebook.com,intouchsol.com -t http://www.google.com,,http://www.intouchsol.com')}\n`));
    },
    inputURLsAndTargetRedirectStatusCodesLengthMismatched: () => {
        console.error(red(`\n    ERROR: The number of input URLs and target status codes (sites and codes) did not match.\n    If you use codes, the list of codes must be the same length as the list of sites.\n    To skip a code for a specific site, just leave an empty spot in the code list.\n    For example, to skip providing a code for facebook.com:\n    ${magenta('redirect-tester -s google.com,facebook.com,intouchsol.com -c 301,,301')}\n`));
    },
    targetRedirectStatusCodesTypeInvalid: () => {
        console.error(red(`\n    ERROR: Some target status codes were not numbers. Codes must be three-digit numbers.`));
    },
    failedToParseAuth: () => {
        console.error(red(`\n    ERROR: Could not parse auth option (username and password), please check the format:\n    ${magenta('redirect-tester -s example.com -a username:password')}\n`));
    },
    noReportsBeingCreated: () => {
        console.warn(yellow('\n    WARNING: No reports are being created. Use the --json or --xlsx options to generate reports.'));
    },
    prefixWarning: (missingPrefixURLs: {inputIndex: number, inputURL: string}[]) => {
        console.warn(yellow('\n    WARNING: Some input URLs did not have a protocol, and no prefix was provided:\n'));
        missingPrefixURLs.forEach(({ inputURL, inputIndex }) => console.log(yellow(`    ${inputIndex + 1}. ${inputURL}`)));
    },
    invalidError: (invalidURLs: {inputIndex: number, inputURL: string}[]) => {
        didError = true;
        console.error(red('\n    ERROR: Some input URLs were invalid:\n'));
        invalidURLs.forEach(({ inputURL, inputIndex }) => console.log(red(`    ${inputIndex + 1}. ${inputURL}`)));
    },
    requestError: ({ error, url }: {error?: Error, url: string}, inputIndex: number | undefined) => {
        didError = true;
        console.error(red('\n    ERROR: An error occurred during the URL request:'));
        console.log(red(`    ${inputIndex !== undefined ? `${inputIndex + 1}.` : undefined}`, url));
        console.log(red(`    `, error));
    },
    missingTargetResult: ({ guid, url }: RequestData) => {
        didError = true;
        console.error(red('\n    ERROR: Missing targetResult when searching by GUID. This should not happen.\n    Please contact a developer and provide the following information:\n'));
        console.log(magenta(`    targetResult guid = ${guid}\n`));
        console.log(magenta(`    targetResult url = ${url}\n`));
    },
    writingToDisk: (json?: boolean, xlsx?: boolean) => {
        console.log(white(`\n\n    Generating report types:\n    ---------------------------------------------------\n    ${json ? green('✓') : red('x')} JSON    ${xlsx ? green('✓') : red('x')} XLSX\n\n\n    Processing ${cyan(report.length)} URLs: \n    ---------------------------------------------------`));
        report.forEach(({ url }, index) => console.log(`    ${cyan(`${index + 1}.`)} ${white(url)}`));
        console.log('\n');
    },
    errorWritingToDisk: (error: Error, json?: boolean, xlsx?: boolean) => {
        didError = true;
        console.error(red(`\n    ERROR: Error writing ${json ? 'JSON' : xlsx ? 'XLSX' : ''} report to disk:`, error));
    },
    wroteToDisk: (jsonFilenames: string[], xlsxFilenames: string[], json?: boolean, xlsx?: boolean) => {
        if (json) {
            console.log(white(`\n    Wrote JSON report(s) to disk:\n`));
            jsonFilenames.forEach((filename, index) => console.log(white(`        ${index + 1}. ${cyan(filename)}`)));
            console.log('\n');
        }
        if (xlsx) {
            if (!json) {
                console.log('\n');
            }
            console.log(white(`    Wrote XLSX report(s) to disk:\n`));
            xlsxFilenames.forEach((filename, index) => console.log(white(`        ${index + 1}. ${cyan(filename)}`)));
            console.log('\n');
        }
    },
    programDidError: () => {
        console.error(red('\n    ERROR: At least one error occurred while the tool was running.\n    However, the report was able to be completed.\n    Please review the console for any error messages.\n'));
    },
};

const validateURLs = (inputURLs: string[]): InputURLData[] => {
    let validURLs: InputURLData[] = [];
    let invalidURLs: InputURLData[] = [];

    inputURLs.forEach((inputURL, index) => {
        const urlObject = { inputIndex: index, inputURL };
        if (URL_VALIDATION_REGEX.test(inputURL)) {
            validURLs.push(urlObject);
        }
        else {
            invalidURLs.push(urlObject);
        }
    });

    if (invalidURLs.length) {
        log.invalidError(invalidURLs);
    }

    return validURLs;
};

const generateGuid = (): string => {
    const placeholder = (([ 1e7 ] as any) + -1e3 + -4e3 + -8e3 + -1e11);
    const guid = placeholder.replace(/[018]/g, (character: number) => {
        const randomNumber = (randomFillSync(new Uint8Array(1))[0] & 15) >> (character / 4);
        const randomString = (character ^ randomNumber).toString(16);

        return randomString;
    });

    return guid;
};

const creatInitialReport = (
    validURLs: InputURLData[],
    targetURLs: (string | undefined)[],
    targetRedirectStatusCodes: (number | undefined)[],
    prefix?: string,
    protocol?: string
) => {
    let missingPrefixURLs: InputURLData[] = [];

    const baseURLObjects = validURLs.map(({ inputURL, inputIndex }) => {
        let url: string;

        if (URL_PROTOCOL_REGEX.test(inputURL)) {
            url = new URL(inputURL).href;
        }
        else {
            missingPrefixURLs.push({ inputURL, inputIndex });
            url = new URL((protocol ? protocol : 'https://') + (prefix ? prefix : '') + inputURL).href;
        }

        const targetURL = targetURLs[inputIndex];
        const targetRedirectStatusCode = targetRedirectStatusCodes[inputIndex];

        return {
            guid: generateGuid(),
            url,
            inputURL,
            inputIndex,
            targetURL: targetURL,
            targetRedirectStatusCode: targetRedirectStatusCode,
            responses: [],
        };
    });

    if (missingPrefixURLs.length > 0 && !prefix) {
        log.prefixWarning(missingPrefixURLs);
    }

    report.push(...baseURLObjects);
};

const chunk = <T>(items: T[], batchSize: number) => {
    const length = items.length;
    const chunks: T[][] = [];

    for (let i = 0; i < length; i += batchSize) {
        chunks.push(items.slice(i, i + batchSize));
    }

    return chunks;
};

const performBatchAsyncRequests = (requests: RequestData[], auth?: string) => {
    return Promise.all<ResponseData>(requests.map(({ url, ...rest }) => {
        return new Promise(resolve => {
            try {
                const protocolAdapter = URL_HTTPS_PROTOCOL_REGEX.test(url) ? https : http;
                protocolAdapter.get(
                    url,
                    { headers: { 'User-Agent': USER_AGENT_DEFAULT } },
                    ({ statusCode, headers }) => {
                        if (statusCode === 401 && auth) {
                            protocolAdapter.get(
                                url,
                                {
                                    headers: {
                                        'User-Agent': USER_AGENT_DEFAULT,
                                        'Authorization': 'Basic ' + Buffer.from(auth).toString('base64'),
                                    }
                                },
                                ({ statusCode, headers }) => resolve({ ...rest, url, statusCode, location: headers && headers.location })
                            ).on('error', error => resolve({ ...rest, url, error }));
                        }
                        else {
                            resolve({ ...rest, url, statusCode, location: headers && headers.location });
                        }
                    }
                ).on('error', error => resolve({ ...rest, url, error }));
            }
            catch (error) {
                resolve({ ...rest, url, error });
            }
        });
    }));
};

const batchCheckRedirects = async (requests: RequestData[], numberOfConcurrentRequests: number, auth?: string) => {
    const requestChunks = chunk<RequestData>(requests, numberOfConcurrentRequests);
    const chunkedResults: ResponseData[][] = [];

    for (let chunk of requestChunks) {
        chunkedResults.push(await performBatchAsyncRequests(chunk, auth));
    }

    const results = chunkedResults.reduce((accumulator, value) => accumulator.concat(value), []);

    return results;
};

const updateReport = (results: ResponseData[]) => {
    results.forEach(result => {
        const targetBaseURL = report.find(({ guid }) => guid === result.guid);

        if (result.hasOwnProperty('error')) {
            log.requestError(result, targetBaseURL ? targetBaseURL.inputIndex : undefined);
        }

        if (targetBaseURL) {
            targetBaseURL.responses.push(result);
        }
        else {
            log.missingTargetResult(result);
        }
    });
};

const recursivelyCheckRedirectsAndUpdateReport = async (requests: RequestData[], numberOfConcurrentRequests: number, auth?: string) => {
    const results = await batchCheckRedirects(requests, numberOfConcurrentRequests, auth);
    updateReport(results);

    const redirects = results.filter(({ statusCode, location }) => statusCode && location && statusCode >= 300 && statusCode < 400);

    if (redirects.length) {
        const nextRequests = redirects.map(({ guid, location }) => ({ guid, url: location as string }));
        await recursivelyCheckRedirectsAndUpdateReport(nextRequests, numberOfConcurrentRequests, auth);
    }
};

const finalizeReport = () => {
    report.forEach(({ responses }, index) => {
        const reportItem = report[index];
        const redirectsLength = responses.length;
        reportItem.numberOfRedirects = redirectsLength - 1;
        reportItem.finalStatusCode = responses[redirectsLength - 1].statusCode;
        reportItem.finalURL = responses[redirectsLength - 1].url;
        const redirectResponses = responses.filter(({ statusCode }) => (statusCode && statusCode >= 300 && statusCode < 400));
        reportItem.finalRedirectStatusCode = redirectResponses.length ? redirectResponses[redirectResponses.length - 1].statusCode : '';

        if (reportItem.targetURL) {
            reportItem.targetURLMatched = reportItem.targetURL === reportItem.finalURL;
        }

        if (reportItem.targetRedirectStatusCode) {
            reportItem.targetStatusMatched = reportItem.targetRedirectStatusCode === reportItem.finalRedirectStatusCode;
        }
    });
};

const generateReportFilename = (fileExtension: string) => {
    const today = new Date();
    const timestamp = `${today.getFullYear()}-${today.getMonth()}-${today.getDate()}-${today.getHours()}${today.getMinutes()}${today.getSeconds()}`;
    return `${timestamp}_${REPORT_FILENAME_SUFFIX_DEFAULT}.${fileExtension}`;
};

const createExcelWorkbook = (usedTargetURLs: boolean, usedTargetCodes: boolean, prefix?: string, protocol?: string, auth?: string) => {
    const xl = require('excel4node');
    const workBook = new xl.Workbook();
    const workSheet = workBook.addWorksheet('Sheet 1');
    const headerStyle = workBook.createStyle({ fill: { type: 'pattern', patternType: 'solid', fgColor: '#5595D0' }, font: { color: '#FAFAFA', size: 16, bold: true } });
    const asideStyle = workBook.createStyle({ font: { color: '#030308', size: 16, bold: true }, alignment: { horizontal: 'right' } });
    const valueStyle = workBook.createStyle({ font: { color: '#030308', size: 12 } });
    const textWrapStyle = workBook.createStyle({ alignment: { wrapText: true } });

    const headerStartRow = 6;

    const headerTitlesAndWidths = [
        ['#', 6],
        ['Checked URL', 50],
        ['Target URL', 50],
        ['URL Match', 20],
        ['Final URL', 50],
        ['Target Redirect Status', 30],
        ['Status Match', 20],
        ['Final Redirect Status', 25],
        ['Final Response Status', 25],
        ['Redirect Count', 20],
        ['Input URL', 50]
    ];

    headerTitlesAndWidths.forEach(([ string, width ], index) => {
        workSheet.cell(headerStartRow, index + 1).string(string).style(headerStyle);
        workSheet.column(index + 1).setWidth(width);
    });

    workSheet.cell(1, 2).string('Checked URL Count:').style(asideStyle);
    workSheet.cell(1, 3).string(String(report.length)).style(valueStyle);

    workSheet.cell(2, 2).string('Used Target URLs:').style(asideStyle);
    workSheet.cell(2, 3).string(String(usedTargetURLs)).style(valueStyle);

    workSheet.cell(3, 2).string('Used Target Codes:').style(asideStyle);
    workSheet.cell(3, 3).string(String(usedTargetCodes)).style(valueStyle);

    workSheet.cell(1, 4).string('Used Auth:').style(asideStyle);
    workSheet.cell(1, 5).string(auth ? 'true' : 'false').style(valueStyle);

    workSheet.cell(2, 4).string('Prefix:').style(asideStyle);
    workSheet.cell(2, 5).string(prefix ? prefix : '').style(valueStyle);

    workSheet.cell(3, 4).string('Protocol:').style(asideStyle);
    workSheet.cell(3, 5).string(protocol ? protocol : '').style(valueStyle);

    const headersLength = headerTitlesAndWidths.length;

    report.forEach(({inputIndex, url, targetURL, targetURLMatched, finalURL, targetRedirectStatusCode, targetStatusMatched, finalStatusCode, finalRedirectStatusCode, numberOfRedirects, inputURL, responses}, index) => {
        const rowNumber = index + 1 + headerStartRow;
        workSheet.row(rowNumber).setHeight(60);
        workSheet.cell(rowNumber, 1).string(String(inputIndex + 1)).style(valueStyle);
        workSheet.cell(rowNumber, 2).string(url).style(valueStyle).style(textWrapStyle);
        workSheet.cell(rowNumber, 3).string(String(targetURL !== undefined ? targetURL : '')).style(valueStyle).style(textWrapStyle);
        workSheet.cell(rowNumber, 4).string(String(targetURLMatched !== undefined ? targetURLMatched : '')).style(valueStyle);
        workSheet.cell(rowNumber, 5).string(String(finalURL)).style(valueStyle).style(textWrapStyle);
        workSheet.cell(rowNumber, 6).string(String(targetRedirectStatusCode !== undefined ? targetRedirectStatusCode : '')).style(valueStyle);
        workSheet.cell(rowNumber, 7).string(String(targetStatusMatched !== undefined ? targetStatusMatched : '')).style(valueStyle);
        workSheet.cell(rowNumber, 8).string(String(finalRedirectStatusCode)).style(valueStyle);
        workSheet.cell(rowNumber, 9).string(String(finalStatusCode)).style(valueStyle);
        workSheet.cell(rowNumber, 10).string(String(numberOfRedirects)).style(valueStyle);
        workSheet.cell(rowNumber, 11).string(inputURL).style(valueStyle).style(textWrapStyle);
        
        responses.forEach(({ statusCode, location }, index) => {
            const responseNumber = index + 1;
            const columnNumber = headersLength - 1 + (responseNumber * 2);
            workSheet.column(columnNumber).setWidth(25);
            workSheet.column(columnNumber + 1).setWidth(40);
            workSheet.cell(headerStartRow, columnNumber).string(`Response ${responseNumber} Status`).style(headerStyle);
            workSheet.cell(rowNumber, columnNumber).string(String(statusCode ? statusCode : '')).style(valueStyle);
            workSheet.cell(headerStartRow, columnNumber + 1).string(`Response ${responseNumber} Location`).style(headerStyle);
            workSheet.cell(rowNumber, columnNumber + 1).string(location ? location : '').style(valueStyle).style(textWrapStyle);
        });
    });

    return workBook;
};

const writeToDisk = (
    usedTargetURLs: boolean,
    usedTargetCodes: boolean,
    json?: boolean,
    xlsx?: boolean,
    prefix?: string,
    protocol?: string,
    auth?: string,
    filename?: string,
) => {
    log.writingToDisk(json, xlsx);

    const jsonFilenames: string[] = [];
    const xlsxFilenames: string[] = [];

    if (json) {
        try {
            const jsonReport = {
                prefix: prefix ? prefix : '',
                protocol: protocol ? protocol : '',
                auth: auth ? true : false,
                report: report,
            };
            const jsonFilename = filename ? `${filename}.json` : generateReportFilename('json');
            filesystem.writeFile(jsonFilename, JSON.stringify(jsonReport), 'utf8', error => error ? log.errorWritingToDisk(error, true) : undefined);
            jsonFilenames.push(jsonFilename);
        }
        catch (error) {
            log.errorWritingToDisk(error, true);
        }
    }
    if (xlsx) {
        try {
            const workBook = createExcelWorkbook(usedTargetURLs, usedTargetCodes, prefix, protocol, auth);
            const xlsxFilename = filename ? `${filename}.xlsx` : generateReportFilename('xlsx');
            workBook.write(xlsxFilename);
            xlsxFilenames.push(xlsxFilename);
        }
        catch (error) {
            log.errorWritingToDisk(error, false, true);
        }
    }

    log.wroteToDisk(jsonFilenames, xlsxFilenames, json, xlsx);
};

const init = async () => {
    const { sites: inputURLs, targets: inputTargetURLs, codes: inputTargetRedirectStatusCodes, prefix, protocol, auth, concurrent, json, xlsx, filename } = program;
    const [ username, password ] = auth ? auth.split(':') : [];

    if (inputURLs.length === 0) {
        log.missingInputURLs();
        process.exit(1);
    }

    if (auth && (!username || !password)) {
        log.failedToParseAuth();
        process.exit(1);
    }

    if (inputTargetURLs.length > 0 && inputURLs.length !== inputTargetURLs.length) {
        log.inputAndTargetURLsLengthMismatched();
        process.exit(1);
    }

    if (inputTargetRedirectStatusCodes.length > 0 && inputURLs.length !== inputTargetRedirectStatusCodes.length) {
        log.inputURLsAndTargetRedirectStatusCodesLengthMismatched();
        process.exit(1);
    }

    const targetRedirectStatusCodes = inputTargetRedirectStatusCodes.map((statusCodeString: string) => statusCodeString !== '' ? parseInt(statusCodeString) : undefined);

    if (inputTargetRedirectStatusCodes.length && targetRedirectStatusCodes.some((statusCode: number | undefined) => statusCode !== undefined && isNaN(statusCode))) {
        log.targetRedirectStatusCodesTypeInvalid();
        process.exit(1);
    }

    if (!json && !xlsx) {
        log.noReportsBeingCreated();
    }

    const validURLs = validateURLs(inputURLs);
    creatInitialReport(validURLs, inputTargetURLs, targetRedirectStatusCodes, prefix, protocol);

    const initialRequests = report.map(({ guid, url }) => ({ guid, url }));
    const  concurrentNumber = concurrent ? parseInt(concurrent) : CONCURRENT_REQUESTS_DEFAULT;
    await recursivelyCheckRedirectsAndUpdateReport(initialRequests, concurrentNumber, auth);

    finalizeReport();
    
    writeToDisk(
        Boolean(inputTargetURLs.length),
        Boolean(inputTargetRedirectStatusCodes.length),
        json,
        xlsx,
        prefix,
        protocol,
        auth,
        filename,
    );

    if (didError) {
        log.programDidError();
    }
};

init();