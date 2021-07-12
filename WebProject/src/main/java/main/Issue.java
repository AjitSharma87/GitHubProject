package jira;

import java.io.File;

import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Base64;
import javax.json.*;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.ContentType;
import org.apache.http.entity.mime.MultipartEntityBuilder;
import org.apache.http.entity.mime.content.FileBody;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClientBuilder;
import org.apache.http.HttpEntity;


/**
 * Represents a Jira Issue.  Methods are supplied to edit and add to
 * multiple fields.
 * 
 * @author Ian Polidora
 * @version 1.0
 */
public class Issue {
	private String hostUrl;
	private String authString;
	private String issueIdentifier;
	
	/**
	 * Creates an Issue that is connected to the specified url.  The users credentials
	 * must be passed as well in order to authorize any actions.
	 * 
	 * @param hostUrl The root url of the Jira site that will be accessed.
	 * @param authString A string containing the user's credentials. Must be in the form username:apiKey.
	 * @param issueIdentifier The id of the issue that will be acted on.
	 */
	public Issue(String hostUrl, String authString, String issueIdentifier) {
		this.hostUrl = hostUrl;
		this.authString = authString;
		this.issueIdentifier = issueIdentifier;
	}
	
	/**
	 * Updates the summary field of the specified Jira issue.  
	 * 
	 * @param summary String value that holds the new value of the summary.
	*/
	public void updateSummary(String summary) {
		try {
			// Create the URL for the request.
			URL url = new URL(hostUrl + "/rest/api/2/issue/" + issueIdentifier);
			
			// Create the Http Connection.
			HttpURLConnection connection = JiraClient.establishConnection(url);
			
			// Create the body of the JSON request.
			String update = Json.createObjectBuilder().add("update", Json.createObjectBuilder()
				    .add("summary", Json.createArrayBuilder().add(Json.createObjectBuilder().add("set", summary))))
				    .build().toString();
			
			// Add appropriate request headers.
			connection.setRequestMethod("PUT");
            connection.setRequestProperty("Authorization", "Basic " + Base64.getEncoder().encodeToString(authString.getBytes()));
            connection.setRequestProperty("Content-Type", "application/json");
			
            // Write the JSON body to the request.
            connection.getOutputStream().write(update.getBytes());
			
            // Send request and retrieve the input stream.
            connection.getInputStream();
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Updates the description field of the specified Jira issue.  
	 * 
	 * @param description String value that holds the new value of the description.
	*/
	public void updateDescription(String description) {
		try {
			// Create the URL for the request.
			URL url = new URL(hostUrl + "/rest/api/2/issue/" + issueIdentifier);
			
			// Create the Http Connection.
			HttpURLConnection connection = JiraClient.establishConnection(url);
			
			// Create the body of the JSON request.
			String update = Json.createObjectBuilder().add("update", Json.createObjectBuilder()
				    .add("description", Json.createArrayBuilder().add(Json.createObjectBuilder().add("set", description))))
				    .build().toString();
			
			// Add appropriate request headers.
			connection.setRequestMethod("PUT");
            connection.setRequestProperty("Authorization", "Basic " + Base64.getEncoder().encodeToString(authString.getBytes()));
            connection.setRequestProperty("Content-Type", "application/json");
            
            // Write the JSON body to the request.
			connection.getOutputStream().write(update.getBytes());
			
			// Send request and retrieve the input stream.
			connection.getInputStream();
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Updates the priority field of the specified Jira issue.  The method takes in
	 * the new priority id as an integer.
	 * 
	 * @param priority Integer value that holds the priority id value.
	*/
	public void updatePriority(int priority) {
		try {
			// Create the URL for the request.
			URL url = new URL(hostUrl + "/rest/api/2/issue/" + issueIdentifier);
			
			// Create the Http Connection.
			HttpURLConnection connection = JiraClient.establishConnection(url);
			
			// Create the body of the JSON request.
			String update = Json.createObjectBuilder().add("update", Json.createObjectBuilder()
				    .add("priority", Json.createArrayBuilder().add(Json.createObjectBuilder().add("set", Json.createObjectBuilder().add("id", String.valueOf(priority))))))
				    .build().toString();

			// Add appropriate request headers.
			connection.setRequestMethod("PUT");
            connection.setRequestProperty("Authorization", "Basic " + Base64.getEncoder().encodeToString(authString.getBytes()));
            connection.setRequestProperty("Content-Type", "application/json");
            
            // Write the JSON body to the request.
			connection.getOutputStream().write(update.getBytes());
			
			// Send request and retrieve the input stream.
			connection.getInputStream();
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Adds a comment to the specified Jira issue.  
	 * 
	 * @param comment String value that holds the contents of the comment.
	*/
	public void addComment(String comment) {
		try {
			// Create the URL for the request.
			URL url = new URL(hostUrl + "/rest/api/2/issue/" + issueIdentifier + "/comment");
			
			// Create the Http Connection.
			HttpURLConnection connection = JiraClient.establishConnection(url);
			
			// Create the body of the JSON request.
			String update = Json.createObjectBuilder().add("body", comment).build().toString();

			// Add appropriate request headers.
			connection.setRequestProperty("X-Atlassian-Token", "no-check");
			connection.setRequestMethod("POST");
            connection.setRequestProperty("Authorization", "Basic " + Base64.getEncoder().encodeToString(authString.getBytes()));
            connection.setRequestProperty("Content-Type", "application/json");
			
            // Write the JSON body to the request.
            connection.getOutputStream().write(update.getBytes());
            
            // Send request and retrieve the input stream.
			connection.getInputStream();
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Updates the environment field of the specified Jira issue.  
	 * 
	 * @param environment String value that holds the new value of the environment.
	*/
	public void updateEnvironment(String environment) {
		try {
			// Create the URL for the request.
			URL url = new URL(hostUrl + "/rest/api/2/issue/" + issueIdentifier);
			
			// Create the Http Connection.
			HttpURLConnection connection = JiraClient.establishConnection(url);
			
			// Create the body of the JSON request.
			String update = Json.createObjectBuilder().add("update", Json.createObjectBuilder()
				    .add("environment", Json.createArrayBuilder().add(Json.createObjectBuilder().add("set", environment))))
				    .build().toString();

			// Add appropriate request headers.
			connection.setRequestMethod("PUT");
            connection.setRequestProperty("Authorization", "Basic " + Base64.getEncoder().encodeToString(authString.getBytes()));
            connection.setRequestProperty("Content-Type", "application/json");
            
            // Write the JSON body to the request.
			connection.getOutputStream().write(update.getBytes());
			
			// Send request and retrieve the input stream.
			connection.getInputStream();
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Adds and attachment to the specified Jira issue. A filepath should be specified
	 * for the file or image that should be added. 
	 * 
	 * @param filePath String value that holds the path to the file that will be attached.
	*/
	public void addAttachment(String filePath) {
		try {
			// Create the Http client for sending requests.
			CloseableHttpClient httpclient = HttpClientBuilder.create().build();
			
			// Create a Multipart entity builder to hold the file.
			MultipartEntityBuilder multipartEntity = MultipartEntityBuilder.create();
			
			// Create a POST request object.
			HttpPost httppost = new HttpPost(hostUrl + "/rest/api/2/issue/" + issueIdentifier + "/attachments");
			
			// Add the appropriate headers.
			httppost.setHeader("X-Atlassian-Token", "no-check");
			httppost.setHeader("Authorization", "Basic " + Base64.getEncoder().encodeToString(authString.getBytes()));
			
			// Create a file object from the supplied path.
			File file = new File(filePath);
			
			// Retrieve the file body from the file.
			FileBody fileBody = new FileBody(file, ContentType.MULTIPART_FORM_DATA, file.getName());
			
			// Add the file to the multipart entity.
			multipartEntity.addPart("file", fileBody);
			
			// Create an Http entity from the multipart entity.
			HttpEntity httpEntity = multipartEntity.build();

			// Add the entity to the post request.
			httppost.setEntity(httpEntity);
			
			// Execute the POST request.
			httpclient.execute(httppost);
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Adds and attachment to the specified Jira issue. A filepath should be specified
	 * for the file or image that should be added. 
	 * 
	 * @param file File object that contains the file that will be attached.
	*/
	public void addAttachment(File file) {
		try {
			// Create the Http client for sending requests.
			CloseableHttpClient httpclient = HttpClientBuilder.create().build();
			
			// Create a Multipart entity builder to hold the file.
			MultipartEntityBuilder multipartEntity = MultipartEntityBuilder.create();
			
			// Create a POST request object.
			HttpPost httppost = new HttpPost(hostUrl + "/rest/api/2/issue/" + issueIdentifier + "/attachments");
			
			// Add the appropriate headers.
			httppost.setHeader("X-Atlassian-Token", "no-check");
			httppost.setHeader("Authorization", "Basic " + Base64.getEncoder().encodeToString(authString.getBytes()));
			
			// Retrieve the file body from the given file.
			FileBody fileBody = new FileBody(file, ContentType.MULTIPART_FORM_DATA, file.getName());
			
			// Add the file to the multipart entity.
			multipartEntity.addPart("file", fileBody);
			
			// Create an Http entity from the multipart entity.
			HttpEntity httpEntity = multipartEntity.build();
			
			// Add the entity to the post request.
			httppost.setEntity(httpEntity);
			
			// Execute the POST request.
			httpclient.execute(httppost);
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Adds a new label to the specified Jira issue.  
	 * 
	 * @param labelName String value that holds the label name.
	*/
	public void addLabel(String labelName) {
		try {
			// Create the URL for the request.
			URL url = new URL(hostUrl + "/rest/api/2/issue/" + issueIdentifier);
			
			// Create the Http Connection.
			HttpURLConnection connection = JiraClient.establishConnection(url);
			
			// Create the body of the JSON request.
			String update = Json.createObjectBuilder().add("update", Json.createObjectBuilder()
				    .add("labels", Json.createArrayBuilder().add(Json.createObjectBuilder().add("add", labelName))))
				    .build().toString();

			// Add appropriate request headers.
			connection.setRequestMethod("PUT");
            connection.setRequestProperty("Authorization", "Basic " + Base64.getEncoder().encodeToString(authString.getBytes()));
            connection.setRequestProperty("Content-Type", "application/json");
            
            // Write the JSON body to the request.
			connection.getOutputStream().write(update.getBytes());
			
			// Send request and retrieve the input stream.
			connection.getInputStream();
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Removes a label from the specified Jira issue.  
	 * 
	 * @param labelName String value that holds the name of the label that will be removed.
	*/
	public void removeLabel(String labelName) {
		try {
			// Create the URL for the request.
			URL url = new URL(hostUrl + "/rest/api/2/issue/" + issueIdentifier);
			
			// Create the Http Connection.
			HttpURLConnection connection = JiraClient.establishConnection(url);
			
			// Create the body of the JSON request.
			String update = Json.createObjectBuilder().add("update", Json.createObjectBuilder()
				    .add("labels", Json.createArrayBuilder().add(Json.createObjectBuilder().add("remove", labelName))))
				    .build().toString();

			// Add appropriate request headers.
			connection.setRequestMethod("PUT");
            connection.setRequestProperty("Authorization", "Basic " + Base64.getEncoder().encodeToString(authString.getBytes()));
            connection.setRequestProperty("Content-Type", "application/json");
            
            // Write the JSON body to the request.
			connection.getOutputStream().write(update.getBytes());
			
			// Send request and retrieve the input stream.
			connection.getInputStream();
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	public void changeStatus(int statusId) {
		try {
			// Create the URL for the request.
			URL url = new URL(hostUrl + "/rest/api/2/issue/" + issueIdentifier + "/transitions");
			
			// Create the Http Connection.
			HttpURLConnection connection = JiraClient.establishConnection(url);
			
			// Create the body of the JSON request.
			String update = Json.createObjectBuilder().add("transition", Json.createObjectBuilder()
				    .add("id", String.valueOf(statusId)))
				    .build().toString();

			// Add appropriate request headers.
			connection.setRequestMethod("POST");
            connection.setRequestProperty("Authorization", "Basic " + Base64.getEncoder().encodeToString(authString.getBytes()));
            connection.setRequestProperty("Content-Type", "application/json");
            
            // Write the JSON body to the request.
			connection.getOutputStream().write(update.getBytes());
			
			// Send request and retrieve the input stream.
			connection.getInputStream();
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	public void changeAssignee(String accountId) {
		try {
			// Create the URL for the request.
			URL url = new URL(hostUrl + "/rest/api/2/issue/" + issueIdentifier + "/assignee");
			
			// Create the Http Connection.
			HttpURLConnection connection = JiraClient.establishConnection(url);
			
			// Create the body of the JSON request.
			String update = Json.createObjectBuilder().add("accountId", accountId)
				    .build().toString();

			// Add appropriate request headers.
			connection.setRequestMethod("PUT");
            connection.setRequestProperty("Authorization", "Basic " + Base64.getEncoder().encodeToString(authString.getBytes()));
            connection.setRequestProperty("Content-Type", "application/json");
            
            // Write the JSON body to the request.
			connection.getOutputStream().write(update.getBytes());
			
			// Send request and retrieve the input stream.
			connection.getInputStream();
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
}
