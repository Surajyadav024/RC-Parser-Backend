const express = require('express');
const readXlsxFile = require('read-excel-file/node')
const path = require('path');
const axios = require('axios');
const bodyParser = require('body-parser');
const cors = require('cors');
const qs = require('qs');
const multer = require('multer');
const serverless = require("serverless-http");

const filepath1 = "RCfile.xlsx";
require('dotenv').config({ path: __dirname + '/.env' });

const REFRESH_TOKEN = process.env.REFRESH_TOKEN;
const ZOHO_CLIENT_ID = process.env.ZOHO_CLIENT_ID;
const ZOHO_CLIENT_SECRET = process.env.ZOHO_CLIENT_SECRET;
const ZOHO_REDIRECT_URI = process.env.ZOHO_REDIRECT_URI;
const ZOHO_TOKEN_URL = process.env.ZOHO_TOKEN_URL;

let projectName;
let projectId;

const app = express();
const upload = multer();

// Middleware
app.use(cors());
app.use(bodyParser.json());



app.get('/', (req, res) => {
    res.json({
        "message": "server is up"
    })
})

app.post("/parse-excel", upload.single("file"), async (req, res) => {
    try {
        const buffer = req.file.buffer;
        console.log("📦 Received buffer of size:", buffer.length);
        const accessToken = await generateAccessToken();
        console.log("access token : " + accessToken);
        const rows = await readXlsxFile(buffer);
        const jsonData = parseExcelRowsToJson(rows);
        const allProjectsData = await getAllProjects(accessToken);
        projectId = findProjectIdByName(jsonData, allProjectsData);
        if (!projectId) {
            return res.status(404).json({ error: "Project ID not found for the given project name." });
        }
        const taskData = await getTasksForProject(projectId, accessToken);
        const updatedJson = assignTaskIdsFromZoho(jsonData, taskData);
        const tasklistData = await getTasklistFromZoho(projectId, accessToken);
        const finalJson = updateTasklistIds(tasklistData, updatedJson);
        const updateLogs = await updateMatchedTasksInZohoProjects(finalJson.project_data, accessToken, projectId);

        res.status(200).json({
            message: "Tasks processed.",
            taskData,
            updateLogs,
        });
        // res.status(200).json(finalJson);

    } catch (error) {
        console.error("Error reading Excel file:", error);
        res.status(500).json({ error: "Failed to process Excel file" });
    }
});


app.listen(3000, () => {
    console.log("Server running on http://localhost:3000");
});

// module.exports = app;
// module.exports.handler = serverless(app);



async function generateAccessToken() {
    try {
        const url = `${ZOHO_TOKEN_URL}?grant_type=refresh_token&client_id=${ZOHO_CLIENT_ID}&client_secret=${ZOHO_CLIENT_SECRET}&redirect_uri=${ZOHO_REDIRECT_URI}&refresh_token=${REFRESH_TOKEN}`;
        const response = await axios.post(url);

        if (response.data && response.data.access_token) {
            console.log("access token generated: ", response.data.access_token)
            return response.data.access_token;
        } else {
            throw new Error("Access token not found in the response.");
        }
    } catch (error) {
        console.error("Error fetching access token:", error.message);
        throw error;
    }
}


function parseExcelRowsToJson(rows) {
    const project_name = rows[1][0]; // First row, first column
    projectName = project_name;
    const project_data = [];
    let currentTasklist = null;
    let taskId = 0;

    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const type = row[1]; // Typ
        const scope = row[2];
        const name = row[3]; // Nazwa
        const group = row[4]; // Grupa
        const progress = row[5] * 100; // Postęp [%]
        const target = row[6];
        const deviation = row[7];
        const indicator = row[8];
        const workTime = row[9];
        const workTimeFromInterval = row[10];

        if (type === "etap") {
            if (currentTasklist) {
                project_data.push(currentTasklist);
            }
            taskId = 0;

            const tasklist = {};
            tasklist.id = 0;
            if (name != null) tasklist.name = name;
            if (group != null) tasklist.group = group;
            if (!isNaN(progress)) tasklist.progress = progress;
            if (type != null) tasklist.type = type;
            if (target != null) tasklist.target = target;
            if (deviation != null) tasklist.deviation = deviation;
            if (indicator != null) tasklist.indicator = indicator;
            if (workTime != null) tasklist.workTime = workTime;
            if (workTimeFromInterval != null) tasklist.workTimeFromInterval = workTimeFromInterval;

            currentTasklist = {
                tasklist,
                tasks: []
            };

        } else if ((type === "czynność" || type === "przedmiot") && currentTasklist) {
            const task = {
                id: 0,
                tasklist_id : 0,
                myId: taskId++,
                name: cleanTaskName(name) || "",
            };
            if (!isNaN(progress)) task.percentage = progress;
            if (group != null) task.group = group;
            if (scope != null) task.scope = scope;
            if (type != null) task.type = type;
            if (target != null) task.target = target;
            if (deviation != null) task.deviation = deviation;
            if (indicator != null) task.indicator = indicator;
            if (workTime != null) task.workTime = workTime;
            if (workTimeFromInterval != null) task.workTimeFromInterval = workTimeFromInterval;

            currentTasklist.tasks.push(task);
        }
    }

    if (currentTasklist) {
        project_data.push(currentTasklist);
    }

    return { project_data };
}


function cleanTaskName(name) {
    // console.log("name is : ", name.replace(/^\s*\d+(\.\d+)*\)?\s*/, ''));
    // return name ? name.replace(/^\s*\d+(\.\d+)*\)?\s*/, '') : '';
    if (!name) return '';
    
    // Remove leading numbers with dots (like "10100.044"), optional closing parenthesis, then any pipe or dot or dash followed by space
    const cleaned = name.replace(/^\s*\d+(\.\d+)*\)?\s*[\.\|\-]?\s*/, '');
    
    console.log("Cleaned name:", cleaned);
    return cleaned;
}

//all-projects  https://projectsapi.zoho.com/restapi/portal/sunreefyachts/projects/"

async function getAllProjects(accessToken) {
    try {
        const response = await axios.get("https://projectsapi.zoho.com/restapi/portal/gauravheliostechlabsdotcom/projects/", {
            headers: {
                Authorization: `Zoho-oauthtoken ${accessToken}`
            },
            params: {
                status: "active",               // Fetch active projects
                sort_column: "last_modified_time",
                sort_order: "descending",
                index: 1,
                range: 100
            }
        });
        return response.data;
    } catch (error) {
        console.error("Error fetching projects:", error);
    }
}


function findProjectIdByName(parsedExcelJson, allProjectsData) {

    if (!projectName || !allProjectsData.projects) {
        return null;
    }

    const matchedProject = allProjectsData.projects.find(proj =>
        proj.name?.trim().toLowerCase() === projectName.trim().toLowerCase()
    );

    return matchedProject ? matchedProject.id_string : null;
}


async function getTasksForProject(projectId, accessToken) {
    try {
        console.log("this is projectId : " + projectId);
        const response = await axios.get("https://projectsapi.zoho.com/restapi/portal/gauravheliostechlabsdotcom/projects/" + projectId + "/tasks/", {
            headers: {
                Authorization: `Zoho-oauthtoken ${accessToken}`
            },
        });
        return response.data;
    } catch (error) {
        console.error("Error fetching projects:", error);
    }
}

async function getTasklistFromZoho(projectId, accessToken) {
    try {
        console.log("this is projectId : " + projectId);
        const response = await axios.get("https://projectsapi.zoho.com/restapi/portal/gauravheliostechlabsdotcom/projects/" + projectId + "/tasklists/", {
            headers: {
                Authorization: `Zoho-oauthtoken ${accessToken}`
            },
        });
        return response.data;
    } catch (error) {
        console.error("Error fetching projects:", error);
    }
}


function assignTaskIdsFromZoho(parsedJson, zohoTasks) {
    console.log("inside assigning tasks");
    console.log("length of parsendJson project data: ", parsedJson?.project_data?.length);
    console.log("tasks from zoho : ", zohoTasks.tasks.length);
    if (!parsedJson?.project_data?.length || !zohoTasks.tasks?.length) return parsedJson;

        for (const taskGroup of parsedJson.project_data) {
            for (const task of zohoTasks.tasks){
                if(taskGroup.tasklist.name && task.tasklist.name && task.tasklist.name.trim().toLowerCase() === taskGroup.tasklist.name.trim().toLowerCase()){
                    console.log("taks names: ", taskGroup.tasklist.name, "🟰" , task.tasklist.name);
                    let tasklist_id = task.tasklist.id_string;
                    taskGroup.tasklist.id = tasklist_id;
                    taskGroup.tasks.forEach(task => {
                        task.tasklist_id = tasklist_id;
                    });
                }
            }
        }

        for (const taskGroup of parsedJson.project_data) {
            for (const task of taskGroup.tasks) {
                const match = zohoTasks.tasks.find(apiTask =>
                {
                    return apiTask.name.trim().toLowerCase() === task.name.trim().toLowerCase()
                }
                );
                if (match) {
                    console.log("name matched");
                    task.id = match.id_string; // or match.id if you prefer numeric
                }

            }
        }   

    return parsedJson;
}


const updateMatchedTasksInZohoProjects = async (projectData, accessToken, projectId) => {
    const updateLogs = [];
    const updatePromises = [];

    for (const taskList of projectData) {
        for (const task of taskList.tasks) {
            if (task.id !== 0) {
                // Update existing tasks
                console.log("📩📩sending update request-->");
                const payload = {
                    percent_complete: Math.round(task.percentage / 10) * 10
                };
                  
                const customFields = {};
                  
                if (task.type != null) customFields["UDF_CHAR8"] = task.type;
                if (task.scope != null) customFields["UDF_CHAR9"] = task.scope;
                if (task.group != null) customFields["UDF_CHAR10"] = task.group;
                if (task.target != null) customFields["UDF_CHAR11"] = task.target;
                if (task.deviation != null) customFields["UDF_CHAR12"] = task.deviation;
                if (task.indicator != null) customFields["UDF_CHAR5"] = task.indicator;
                if (task.workTime != null) customFields["UDF_CHAR6"] = task.workTime;
                if (task.workTimeFromInterval != null) customFields["UDF_CHAR7"] = task.workTimeFromInterval;
                  
                if (Object.keys(customFields).length > 0) {
                    payload.custom_fields = JSON.stringify(customFields);
                }
                  
                const options = {
                    method: "POST",
                    url: `https://projectsapi.zoho.com/restapi/portal/gauravheliostechlabsdotcom/projects/${projectId}/tasks/${task.id}/`,
                    headers: {
                        Authorization: `Zoho-oauthtoken ${accessToken}`,
                        "Content-Type": "application/x-www-form-urlencoded"
                    },
                    data: qs.stringify(payload)
                };

                updatePromises.push(
                    axios(options)
                        .then((response) => ({
                            status: "fulfilled",
                            taskName: task.name,
                            id: task.id,
                            responseData: response.data
                        }))
                        .catch((error) => ({
                            status: "rejected",
                            taskName: task.name,
                            id: task.id,
                            error: error.response ? error.response.data : error.message
                        }))
                );
            } else if (task.id == 0 && task.tasklist_id != 0) {
                updatePromises.push(handleCreateAndUpdate(task, accessToken, projectId));
            }
        }
    }

    const results = await Promise.allSettled(updatePromises);

    for (const result of results) {
        if (result.status === "fulfilled") {
            const { taskName, id, responseData } = result.value;
            updateLogs.push({
                status: "✅ SUCCESS",
                taskName,
                id,
                response: responseData
            });
        } else {
            const { taskName, id, error } = result.reason || result.value;
            updateLogs.push({
                status: "❌ FAILED",
                taskName,
                id,
                error
            });
        }
    }

    return updateLogs;
};

const handleCreateAndUpdate = async (task, accessToken, projectId) => {
    try {
        console.log("sending create request🔉🔉🔉")
        // Create the task
        const createPayload = {
            name: task.name, // Use the actual task name instead of hardcoded value
            tasklist_id: task.tasklist_id,
        };

        const customFields = {};
        
        if (task.type != null) customFields["UDF_CHAR8"] = task.type;
        if (task.scope != null) customFields["UDF_CHAR9"] = task.scope;
        if (task.group != null) customFields["UDF_CHAR10"] = task.group;
        if (task.target != null) customFields["UDF_CHAR11"] = task.target;
        if (task.deviation != null) customFields["UDF_CHAR12"] = task.deviation;
        if (task.indicator != null) customFields["UDF_CHAR5"] = task.indicator;
        if (task.workTime != null) customFields["UDF_CHAR6"] = task.workTime;
        if (task.workTimeFromInterval != null) customFields["UDF_CHAR7"] = task.workTimeFromInterval;
        
        if (Object.keys(customFields).length > 0) {
            createPayload.custom_fields = JSON.stringify(customFields);
        }
        
        const createOptions = {
            method: "POST",
            url: `https://projectsapi.zoho.com/restapi/portal/gauravheliostechlabsdotcom/projects/${projectId}/tasks/`,
            headers: {
                Authorization: `Zoho-oauthtoken ${accessToken}`,
                "Content-Type": "application/x-www-form-urlencoded"
            },
            data: qs.stringify(createPayload)
        };

        const createResp = await axios(createOptions);
        const taskId = createResp.data.tasks[0].id_string;
        console.log("this is taskId", taskId);
        
        // Create update payload and options
        const updatePayload = {
            percent_complete: Math.round(task.percentage / 10) * 10
        };
        
        const updateOptions = {
            method: "POST",
            url: `https://projectsapi.zoho.com/restapi/portal/gauravheliostechlabsdotcom/projects/${projectId}/tasks/${taskId}/`,
            headers: {
                Authorization: `Zoho-oauthtoken ${accessToken}`,
                "Content-Type": "application/x-www-form-urlencoded"
            },
            data: qs.stringify(updatePayload)
        };
        
        // Execute the update request
        const updateResp = await axios(updateOptions);
        
        return {
            status: "fulfilled",
            taskName: task.name,
            id: taskId,
            responseData: {
                create: createResp.data,
                update: updateResp.data
            }
        };
    } catch (error) {
        return {
            status: "rejected",
            taskName: task.name,
            id: "creation_failed",
            error: error.response ? error.response.data : error.message
        };
    }
};


function updateTasklistIds(tasklistData, updatedJson){
    tasklistData.tasklists.forEach(tasklist => {
        const zohoTasklistName = tasklist.name.trim().toLowerCase();
        const zohoTasklistid = tasklist.id_string;

        updatedJson.project_data.forEach(taskGroup => {
            const tasklistName = taskGroup.tasklist.name.trim().toLowerCase();
            if(zohoTasklistName === tasklistName){
                taskGroup.tasklist.id = zohoTasklistid;
                taskGroup.tasks.forEach(task => {
                    task.tasklist_id = zohoTasklistid;
                })
            }
        })
    })

    return updatedJson;
}
