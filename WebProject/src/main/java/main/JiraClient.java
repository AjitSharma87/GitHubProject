package jira;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Base64;
import javax.json.*;
import org.json.JSONObject;

/**
 * Jira Client provides an interface to interact with Jira Software.  Methods are 
 * supplied to interact with the software and create issues.
 * 
 * @author Ian Polidora
 * @version 1.0
 */
public class JiraClient {
	private String authString;
	private String hostUrl;
	
	/**
	 * Creates a JiraClient that is connected to the specified url.  The users credentials
	 * must be passed as well in order to authorize any actions.
	 * 
	 * @param username The user's username.
	 * @param apiKey The apiKey associated with the account.
	 * @param hostUrl The root url of the Jira site that will be accessed.
	 */
	public JiraClient(String username, String apiKey, String hostUrl) {
		authString = username + ":" + apiKey;
		this.hostUrl = hostUrl;
	}

	/**
	 * Sets the authString field to the new value.
	 * 
	 * @param username New username value.
	 * @param apiKey New API Key value.
	 */
	public void setAuthString(String username, String apiKey) {
		this.authString = username + ":" + apiKey;
	}

	/**
	 * Returns the host url associated with the client object.
	 * 
	 * @return The host url.
	 */
	public String getHostUrl() {
		return hostUrl;
	}

	/**
	 * Set the host url field to the new value.
	 * 
	 * @param hostUrl The new host url.
	 */
	public void setHostUrl(String hostUrl) {
		this.hostUrl = hostUrl;
	}
	
	/**
	 * Creates the HttpURLConnection and does initial set up of the necessary parameters.
	 * 
	 * @param url The URL object associated to the connection.
	 * @return A HttpURLConnection object.
	 */
	static HttpURLConnection establishConnection(URL url) {
		try {
			// Create HTTP connection to manage requests
	        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
	        
	        // Setup initial parameters.
	        connection.setDoOutput(true);
	        connection.setDoInput(true);
	        
	        // Return the HttpURLConnection object.
	        return connection;
		} catch(Exception e) {
			e.printStackTrace();
			return null;
		}
	}
	
	/**
	 * Creates a response object from the supplied InputStream.  
	 * 
	 * @param input InputStream object that contains the response information.
	 * @return A JSONObject that holds the response.
	 */
	static JSONObject getResponseObject(InputStream input) {
		// Initialize String variables.
		String line;
		StringBuffer responseString = new StringBuffer();
		
		// Create a reader from the InputStream.
		BufferedReader buffReader = new BufferedReader(new InputStreamReader(input));

		try {
			// Iterate through the response and retrieve the information.
			while ((line = buffReader.readLine()) != null) {
				responseString.append(line);
			}
			// Close the reader.
			buffReader.close();
		} catch(IOException e) {
			e.printStackTrace();
		}
		
		// Return the JSONObject with the response.
		return new JSONObject(responseString.toString());
	}
	
	/**
	 * Create a Jira issue with the specified information under the supplied project. 
	 * 
	 * @param projectKey The project key that the issue will be placed under.
	 * @param issueType The issue type name.
	 * @param summary The summary of the issue.
	 * @return An Issue object.
	 */
	public Issue createIssue(String projectKey, String issueType, String summary) {
		try {
			// Create the URL for the request.
			URL url = new URL(hostUrl + "/rest/api/2/issue");

			// Create the Http Connection.
            HttpURLConnection connection = establishConnection(url);
            
            // Create the body of the JSON request.
            String issue = Json.createObjectBuilder().add("fields",Json.createObjectBuilder()
            						.add("project",Json.createObjectBuilder().add("key", projectKey))
                                    .add("summary", summary)
                                    .add("issuetype",Json.createObjectBuilder().add("name", issueType))).build().toString();

            // Add appropriate request headers.
            connection.setRequestMethod("POST");
            connection.setRequestProperty("Authorization", "Basic " + Base64.getEncoder().encodeToString(authString.getBytes()));
            connection.setRequestProperty("Content-Type", "application/json");
            
            // Write the JSON body to the request.
            connection.getOutputStream().write(issue.getBytes());

            // Return an Issue object.
            return new Issue(hostUrl, authString, getResponseObject(connection.getInputStream()).getString("id"));
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
	} 
	
	/**
	 * Create a Jira issue with the specified information under the supplied project. 
	 * 
	 * @param projectKey The project key that the issue will be placed under.
	 * @param issueType The issue type id.
	 * @param summary The summary of the issue.
	 * @return An Issue object.
	 */
	public Issue createIssue(String projectKey, int issueType, String summary) {
		try {
			// Create the URL for the request.
			URL url = new URL(hostUrl + "/rest/api/2/issue");

			// Create the Http Connection.
            HttpURLConnection connection = establishConnection(url);
			
            // Create the body of the JSON request.
			String issue = Json.createObjectBuilder().add("fields", Json.createObjectBuilder()
								    .add("project", Json.createObjectBuilder().add("key", projectKey))
								    .add("summary", summary)
								    .add("issuetype", Json.createObjectBuilder().add("id", String.valueOf(issueType)))).build().toString();
			
			// Add appropriate request headers.
			connection.setRequestMethod("POST");
	        connection.setRequestProperty("Authorization", "Basic " + Base64.getEncoder().encodeToString(authString.getBytes()));
	        connection.setRequestProperty("Content-Type", "application/json");
			
	        // Write the JSON body to the request.
	        connection.getOutputStream().write(issue.toString().getBytes());
			
	        // Return an Issue object.
			return new Issue(hostUrl, authString, getResponseObject(connection.getInputStream()).getString("id"));
		} catch(Exception e) {
			e.printStackTrace();
			return null;
		}
	}
	
	/**
	 * Create a Jira issue with the specified information under the supplied project. 
	 * 
	 * @param projectKey The project key that the issue will be placed under.
	 * @param issueType The issue type name.
	 * @param summary The summary of the issue.
	 * @param description The description of the issue.
	 * @return An Issue object.
	 */
	public Issue createIssue(String projectKey, String issueType, String summary, String description) {
		try {
			// Create the URL for the request.
			URL url = new URL(hostUrl + "/rest/api/2/issue");

			// Create the Http Connection.
            HttpURLConnection connection = establishConnection(url);
			
            // Create the body of the JSON request.
			String issue = Json.createObjectBuilder().add("fields", Json.createObjectBuilder()
								    .add("project", Json.createObjectBuilder().add("key", projectKey))
								    .add("summary", summary)
								    .add("description", description)
								    .add("issuetype", Json.createObjectBuilder().add("name", issueType))).build().toString();
			
			// Add appropriate request headers.
			connection.setRequestMethod("POST");
	        connection.setRequestProperty("Authorization", "Basic " + Base64.getEncoder().encodeToString(authString.getBytes()));
	        connection.setRequestProperty("Content-Type", "application/json");
			
	        // Write the JSON body to the request.
	        connection.getOutputStream().write(issue.toString().getBytes());
			
	        // Return an Issue object.
			return new Issue(hostUrl, authString, getResponseObject(connection.getInputStream()).getString("id"));
		} catch(Exception e) {
			e.printStackTrace();
			return null;
		}
	}
	
	/**
	 * Create a Jira issue with the specified information under the supplied project. 
	 * 
	 * @param projectKey The project key that the issue will be placed under.
	 * @param issueType The issue type id.
	 * @param summary The summary of the issue.
	 * @param description The description of the issue.
	 * @return An Issue object.
	 */
	public Issue createIssue(String projectKey, int issueType, String summary, String description) {
		try {
			// Create the URL for the request.
			URL url = new URL(hostUrl + "/rest/api/2/issue");

			// Create the Http Connection.
            HttpURLConnection connection = establishConnection(url);
			
            // Create the body of the JSON request.
            String issue = Json.createObjectBuilder().add("fields", Json.createObjectBuilder()
								    .add("project", Json.createObjectBuilder().add("key", projectKey))
								    .add("summary", summary)
								    .add("description", description)
								    .add("issuetype", Json.createObjectBuilder().add("id", String.valueOf(issueType))))
								    .build().toString();
			
            // Add appropriate request headers.
            connection.setRequestMethod("POST");
	        connection.setRequestProperty("Authorization", "Basic " + Base64.getEncoder().encodeToString(authString.getBytes()));
	        connection.setRequestProperty("Content-Type", "application/json");
			
	        // Write the JSON body to the request.
	        connection.getOutputStream().write(issue.toString().getBytes());
			
	        // Return an Issue object.
			return new Issue(hostUrl, authString, getResponseObject(connection.getInputStream()).getString("id"));
		} catch(Exception e) {
			e.printStackTrace();
			return null;
		}
	}
	
	/**
	 * Create a Jira issue with the specified information under the supplied project. 
	 * 
	 * @param projectKey The project key that the issue will be placed under.
	 * @param issueType The issue type name.
	 * @param summary The summary of the issue.
	 * @param description the description of the issue.
	 * @param priority The priority id of the issue.
	 * @return An Issue object.
	 */
	public Issue createIssue(String projectKey, String issueType, String summary, String description, int priority) {
		try {
			// Create the URL for the request.
			URL url = new URL(hostUrl + "/rest/api/2/issue");

			// Create the Http Connection.
            HttpURLConnection connection = establishConnection(url);
			
            // Create the body of the JSON request.
            String issue = Json.createObjectBuilder().add("fields", Json.createObjectBuilder()
								    .add("project", Json.createObjectBuilder().add("key", projectKey))
								    .add("summary", summary)
								    .add("description", description)
								    .add("issuetype", Json.createObjectBuilder().add("name", issueType))
								    .add("priority", Json.createObjectBuilder().add("id", String.valueOf(priority)))).build().toString();
			
            // Add appropriate request headers.
            connection.setRequestMethod("POST");
	        connection.setRequestProperty("Authorization", "Basic " + Base64.getEncoder().encodeToString(authString.getBytes()));
	        connection.setRequestProperty("Content-Type", "application/json");
			
	        // Write the JSON body to the request.
	        connection.getOutputStream().write(issue.toString().getBytes());
			
	        // Return an Issue object.
			return new Issue(hostUrl, authString, getResponseObject(connection.getInputStream()).getString("id"));
		} catch(Exception e) {
			e.printStackTrace();
			return null;
		}
	}
	
	/**
	 * Create a Jira issue with the specified information under the supplied project. 
	 * 
	 * @param projectKey The project key that the issue will be placed under.
	 * @param issueType The issue type id.
	 * @param summary The summary of the issue.
	 * @param description the description of the issue.
	 * @param priority the priority id of the issue.
	 * @return An Issue object.
	 */
	public Issue createIssue(String projectKey, int issueType, String summary, String description, int priority) {
		try {
			// Create the URL for the request.
			URL url = new URL(hostUrl + "/rest/api/2/issue");

			// Create the Http Connection.
            HttpURLConnection connection = establishConnection(url);
			
            // Create the body of the JSON request.
            String issue = Json.createObjectBuilder().add("fields", Json.createObjectBuilder()
								    .add("project", Json.createObjectBuilder().add("key", projectKey))
								    .add("summary", summary)
								    .add("description", description)
								    .add("issuetype", Json.createObjectBuilder().add("id", String.valueOf(issueType)))
									.add("priority", Json.createObjectBuilder().add("id", String.valueOf(priority)))).build().toString();
			
            // Add appropriate request headers.
            connection.setRequestMethod("POST");
	        connection.setRequestProperty("Authorization", "Basic " + Base64.getEncoder().encodeToString(authString.getBytes()));
	        connection.setRequestProperty("Content-Type", "application/json");
			
	        // Write the JSON body to the request.
	        connection.getOutputStream().write(issue.toString().getBytes());
			
	        // Return an Issue object.
			return new Issue(hostUrl, authString, getResponseObject(connection.getInputStream()).getString("id"));
		} catch(Exception e) {
			e.printStackTrace();
			return null;
		}
	}
	
	/**
	 * Create a Jira issue with the specified information under the supplied project. 
	 * 
	 * @param projectKey The project key that the issue will be placed under.
	 * @param issueType The issue type name.
	 * @param summary The summary of the issue.
	 * @param priority the priority id of the issue.
	 * @return An Issue object.
	 */
	public Issue createIssue(String projectKey, String issueType, String summary, int priority) {
		try {
			// Create the URL for the request.
			URL url = new URL(hostUrl + "/rest/api/2/issue");

			// Create the Http Connection.
            HttpURLConnection connection = establishConnection(url);
			
            // Create the body of the JSON request.
            String issue = Json.createObjectBuilder().add("fields", Json.createObjectBuilder()
								    .add("project", Json.createObjectBuilder().add("key", projectKey))
								    .add("summary", summary)
								    .add("issuetype", Json.createObjectBuilder().add("name", issueType))
								    .add("priority", Json.createObjectBuilder().add("id", String.valueOf(priority)))).build().toString();
			
            // Add appropriate request headers.
            connection.setRequestMethod("POST");
	        connection.setRequestProperty("Authorization", "Basic " + Base64.getEncoder().encodeToString(authString.getBytes()));
	        connection.setRequestProperty("Content-Type", "application/json");
			
	        // Write the JSON body to the request.
	        connection.getOutputStream().write(issue.toString().getBytes());
			
	        // Return an Issue object.
			return new Issue(hostUrl, authString, getResponseObject(connection.getInputStream()).getString("id"));
		} catch(Exception e) {
			e.printStackTrace();
			return null;
		}
	}
	
	/**
	 * Create a Jira issue with the specified information under the supplied project. 
	 * 
	 * @param projectKey The project key that the issue will be placed under.
	 * @param issueType The issue type id.
	 * @param summary The summary of the issue.
	 * @param priority The priority id of the issue.
	 * @return An Issue object.
	 */
	public Issue createIssue(String projectKey, int issueType, String summary, int priority) {
		try {
			// Create the URL for the request.
			URL url = new URL(hostUrl + "/rest/api/2/issue");

			// Create the Http Connection.
            HttpURLConnection connection = establishConnection(url);
			
            // Create the body of the JSON request.
            String issue = Json.createObjectBuilder().add("fields", Json.createObjectBuilder()
								    .add("project", Json.createObjectBuilder().add("key", projectKey))
								    .add("summary", summary)
								    .add("issuetype", Json.createObjectBuilder().add("id", String.valueOf(issueType)))
									.add("priority", Json.createObjectBuilder().add("id", String.valueOf(priority)))).build().toString();
			
            // Add appropriate request headers.
            connection.setRequestMethod("POST");
	        connection.setRequestProperty("Authorization", "Basic " + Base64.getEncoder().encodeToString(authString.getBytes()));
	        connection.setRequestProperty("Content-Type", "application/json");
			
	        // Write the JSON body to the request.
	        connection.getOutputStream().write(issue.toString().getBytes());
			
	        // Return an Issue object.
			return new Issue(hostUrl, authString, getResponseObject(connection.getInputStream()).getString("id"));
		} catch(Exception e) {
			e.printStackTrace();
			return null;
		}
	}
	
	/**
	 * Retrieves an issue from a Jira project and returns the Issue object.
	 * 
	 * @param issueId The id of the issue that will be retrieved.
	 * @return An Issue object.
	 */
	public Issue getIssue(int issueId) {
            
            return new Issue(hostUrl, authString, String.valueOf(issueId));
	}
	
	/**
	 * Retrieves an issue from a Jira project and returns the Issue object.
	 * 
	 * @param issueId The key of the issue that will be retrieved.
	 * @return An Issue object.
	 */
	public Issue getIssue(String issueKey) {
            
            return new Issue(hostUrl, authString, issueKey);
	}
	
}
