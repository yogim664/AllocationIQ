

export interface AiResponse {
  response: string;
}

const Apikey = ""

 const endpoint =
    "https://agentleagueproject-resource.services.ai.azure.com/api/projects/agentleagueproject";
  const agentName = "ProjectAnalysisAgent";
  const agentVersion = "3";




export const uploadPdfAndAnalyze = async (
  fileOrText: File | string
): Promise<AiResponse> => {
  console.log("Extracting and analyzing file/text:", fileOrText);

  try {
    const content = await fileOrText;

    const response = await fetch(
      `${endpoint}/openai/v1/responses`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${Apikey}`,
                "Content-Type": "application/json",
        },
        body: JSON.stringify({
          input: [
            {
              role: "user",
              content,
            },
          ],
          agent_reference: {
            name: agentName,
            version: agentVersion,
            type: "agent_reference",
          },
        }),
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Agent execution failed: ${response.status} - ${errorText}`);
    }


    const responseData = await response.json();
    const text = responseData.output[0].content[0].text;


    console.log("text",text)
    return {
      response: text
    };
  } catch (error) {
    console.error("AI analysis failed:", error);
    throw error;
  }
};



export const StaffRecommandAgent = async (
  skills: string[],
  ProjectData: any
) => {

const endpoint = "https://agentleagueproject-resource.services.ai.azure.com/api/projects/agentleagueproject";
const agentName = "TeamAllocation";
const agentVersion = "24";

  try {


    const response = await fetch(
      `${endpoint}/openai/v1/responses`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${Apikey}`,
                 "Content-Type": "application/json",
        },
        body: JSON.stringify({
          input: [
            {
              role: "user",
              content: `${ProjectData} and required skills are ${skills}`,
            },
          ],
          agent_reference: {
            name: agentName,
            version: agentVersion,
            type: "agent_reference",
          },
        }),
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Agent execution failed: ${response.status} - ${errorText}`);
    }


    const responseData = await response.json();
    const text = responseData.output[0].content[0].text;


    console.log("text",text)
    return {
      response: text
    };
  } catch (error) {
    console.error("AI analysis failed:", error);
    throw error;
  }
 


};



export const ProjectPlanAgent = async (
  Staff: any,
  ProjectData: any
): Promise<any[]> => {
  const planEndpoint =
    "https://agentleagueproject-resource.services.ai.azure.com/api/projects/agentleagueproject";
  const planAgentName = "PlanningAgent";
  const planAgentVersion = "8";

  try {
    const response = await fetch(`${planEndpoint}/openai/v1/responses`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${Apikey}`,
             "Content-Type": "application/json",
      },
      body: JSON.stringify({
        input: [
          {
            role: "user",
            content: `Project details: ${JSON.stringify(ProjectData)} and staff details are ${JSON.stringify(Staff)}`,
          },
        ],
        agent_reference: {
          name: planAgentName,
          version: planAgentVersion,
          type: "agent_reference",
        },
      }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Agent execution failed: ${response.status} - ${errorText}`);
    }

    const responseData = await response.json();
    const text = responseData.output[0].content[0].text;

    console.log("text", text);

    const jsonMatch = text.match(/```(?:json)?\s*([\s\S]*?)```/);
    const jsonText = (jsonMatch ? jsonMatch[1] : text).trim();
    const parsed = JSON.parse(jsonText);

    if (Array.isArray(parsed)) {
      return parsed;
    }
    if (parsed?.planItems && Array.isArray(parsed.planItems)) {
      return parsed.planItems;
    }
    if (parsed?.items && Array.isArray(parsed.items)) {
      return parsed.items;
    }

    throw new Error("ProjectPlanAgent response is not a plan array");
  } catch (error) {
    console.error("Project plan agent failed:", error);
    throw error;
  }
};