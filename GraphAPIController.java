package com.example.graphwebhook;

import java.io.ByteArrayInputStream;
import java.text.ParseException;
import java.time.Instant;

//Copyright (c) Microsoft Corporation. All rights reserved.
//Licensed under the MIT License.

import java.time.OffsetDateTime;
import java.time.Period;
import java.time.temporal.TemporalAmount;
import java.util.UUID;
import java.util.concurrent.CompletableFuture;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.ResponseEntity;
import org.springframework.security.core.Authentication;
import org.springframework.security.oauth2.client.OAuth2AuthorizedClient;
import org.springframework.security.oauth2.client.OAuth2AuthorizedClientService;
import org.springframework.security.oauth2.client.annotation.RegisteredOAuth2AuthorizedClient;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import com.azure.core.models.CloudEvent;
import com.azure.core.util.BinaryData;
import com.microsoft.graph.logger.DefaultLogger;
import com.microsoft.graph.models.ChangeNotification;
import com.microsoft.graph.models.ChangeType;
import com.microsoft.graph.models.Subscription;
import com.microsoft.graph.serializer.DefaultSerializer;
import com.microsoft.graph.serializer.OffsetDateTimeSerializer;

/**
 * <p>
 * Sample class that contains controller methods to create, delete, and renew a 
 * Microsoft Graph API subscription that sends events to an Azure Event Grid 
 * partner topic.
 * 
 * This code is meant to used along with full-blown sample in https://github.com/microsoftgraph/java-spring-webhooks-sample
 * </p>
 */
@Controller
public class GraphAPIController {

 private static final String CREATE_SUBSCRIPTION_ERROR = "Error creating subscription";
 private static final String REDIRECT_HOME = "redirect:/";
 private static final String REDIRECT_LOGOUT = "redirect:/logout";

 private final Logger log = LoggerFactory.getLogger(this.getClass());

 @Autowired
 private SubscriptionStoreService subscriptionStore;

// @Autowired
// private CertificateStoreService certificateStore;

 @Autowired
 private OAuth2AuthorizedClientService authorizedClientService;

 @Value("${notifications.host}")
 private String notificationHost;


 /**
  * Subscribes for all events related to changes to profile information in Microsoft Entra ID about 
  * the current ("me") user.
  * @param model the model provided by Spring
  * @param authentication authentication information for the request
  * @param redirectAttributes redirect attributes provided by Spring
  * @param oauthClient a delegated auth OAuth2 client for the authenticated user
  * @return the name of the template used to render the response
  */
 @GetMapping("/subscribe")
 public CompletableFuture<String> delegatedUser(Model model, Authentication authentication,
         RedirectAttributes redirectAttributes,
         @RegisteredOAuth2AuthorizedClient("graph") OAuth2AuthorizedClient oauthClient) {

     final var graphClient = GraphClientHelper.getGraphClient(oauthClient);
     
     // Get the authenticated user's info. See https://docs.microsoft.com/en-us/graph/sdks/create-requests?tabs=java#use-select-to-control-the-properties-returned
     // https://docs.microsoft.com/en-us/graph/query-parameters#select-parameter
     // and https://docs.microsoft.com/en-us/graph/api/resources/users?view=graph-rest-1.0#common-properties
      final var userFuture = graphClient.me().buildRequest()
              .select("displayName,jobTitle,userPrincipalName").getAsync();


     // Create the subscription
     final var subscriptionRequest = new Subscription();
     subscriptionRequest.changeType = ChangeType.UPDATED.toString();
     // Use the following for event notifications via Event Grid's Partner Topic
     subscriptionRequest.notificationUrl = "EventGrid:?azuresubscriptionid=8A8A8A8A-4B4B-4C4C-4D4D-12E12E12E12E&resourcegroup=yourResourceGroup&partnertopic=yourPartnerTopic&location=theNameOfAzureRegionFortheTopic";
     subscriptionRequest.lifecycleNotificationUrl = "EventGrid:?azuresubscriptionid=8A8A8A8A-4B4B-4C4C-4D4D-12E12E12E12E&resourcegroup=yourResourceGroup&partnertopic=yourPartnerTopic&location=theNameOfAzureRegionFortheTopic";
     subscriptionRequest.resource = "me";
     subscriptionRequest.clientState = UUID.randomUUID().toString();
     // subscriptionRequest.includeResourceData = false;
     subscriptionRequest.expirationDateTime = OffsetDateTime.now().plusSeconds(3600);
     final var subscriptionFuture =
             graphClient.subscriptions().buildRequest().postAsync(subscriptionRequest);

     return userFuture.thenCombine(subscriptionFuture, (user, subscription) -> {
         log.info("*** Created subscription {} for user {}", subscription.id, user.displayName);

         // Save the authorized client so we can use it later from the notification controller
         authorizedClientService.saveAuthorizedClient(oauthClient, authentication);

         // Add information to the model
         model.addAttribute("user", user);
         model.addAttribute("subscriptionId", subscription.id);

         final var subscriptionJson =
                 graphClient.getHttpProvider().getSerializer().serializeObject(subscription);
         model.addAttribute("subscription", subscriptionJson);

         // Add record in subscription store
         subscriptionStore.addSubscription(subscription, authentication.getName());

         model.addAttribute("success", "Subscription created.");

         return "delegatedUser";
     }).exceptionally(e -> {
         log.error(CREATE_SUBSCRIPTION_ERROR, e);
         redirectAttributes.addFlashAttribute("error", CREATE_SUBSCRIPTION_ERROR);
         redirectAttributes.addFlashAttribute("debug", e.getMessage());
         return REDIRECT_HOME;
     });
 }

 @PostMapping("/graphApiSubscriptionLifecycleEvents")
 public CompletableFuture<ResponseEntity<String>> handleLifecycleEvent(@RequestBody CloudEvent lifecycleCloudEvent) throws ParseException {
 	
 	log.info("***** Received lifecycle Event with ID and event type: " + 
 	lifecycleCloudEvent.getId() + ", " + lifecycleCloudEvent.getType() + ".");
 	
 	BinaryData eventData = lifecycleCloudEvent.getData();
 	
 	byte [] eventBytes = eventData.toBytes();
 	ByteArrayInputStream eventNotificationDataInputStream = new ByteArrayInputStream(eventBytes);
 	
 	
 	final var serializer = new DefaultSerializer(new DefaultLogger());
     final var notification =
             serializer.deserializeObject(eventNotificationDataInputStream, ChangeNotification.class);
 	
 	 // Look up subscription in store
     var subscription =
             subscriptionStore.getSubscription(notification.subscriptionId.toString());
     log.info("***** Subscription id of the received notification: " + notification.subscriptionId.toString());
     log.info("***** Lifecycle event type received: " + lifecycleCloudEvent.getType());
     
     if(subscription != null) {
         // Get the authorized OAuth2 client for the relevant user
         final var oauthClient =
                 authorizedClientService.loadAuthorizedClient("graph", subscription.userId);

         final var graphClient = GraphClientHelper.getGraphClient(oauthClient);
         Subscription renewedSbscription = new Subscription();
         
         TemporalAmount threeMonths = Period.ofMonths(3); 
         String nowPlustThreeMonths = Instant.now().plus(threeMonths).toString();
         log.info("**** renewing Graph API subscription with new expiraction time of " + nowPlustThreeMonths);
         renewedSbscription.expirationDateTime = OffsetDateTimeSerializer.deserialize(nowPlustThreeMonths);

         graphClient.subscriptions(subscription.subscriptionId)
         	.buildRequest()
         	.patch(renewedSbscription);

         log.info("**** Graph API subscription renewed");
     	
     }
 	return CompletableFuture.completedFuture(ResponseEntity.ok().body(""));
 }
 
 
 /**
  * Deletes a subscription and logs the user out
  * @param subscriptionId the subscription ID to delete
  * @param oauthClient a delegated auth OAuth2 client for the authenticated user
  * @return a redirect to the logout page
  */
 @GetMapping("/unsubscribe")
 public CompletableFuture<String> unsubscribe(
         @RequestParam(value = "subscriptionId") final String subscriptionId,
         @RegisteredOAuth2AuthorizedClient("graph") OAuth2AuthorizedClient oauthClient) {

     final var graphClient = GraphClientHelper.getGraphClient(oauthClient);

     return graphClient.subscriptions(subscriptionId).buildRequest().deleteAsync()
             .thenApply(sub -> {
                 // Remove subscription from store
                 subscriptionStore.deleteSubscription(subscriptionId);

                 // Logout user
                 return REDIRECT_LOGOUT;
             });
 }
	
}