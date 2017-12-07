/*!
 * rb.workflowtools v1.0.10
 * Docs & License: tba
 * (c) 2015 Riu Baring
 */

(function(factory) {
    "use strict";
    if (typeof define === "function" && define.amd) {
        define([ "jquery" ], factory);
    } else {
        factory(jQuery);
    }
})(function($) {
	$.fn.workflowTools = function(options) {
		var theContainer = this;
		var settings = $.extend({
			buttonCollection: {'0': [],'1': [],'2': ['Resume', 'Terminate'],'3': [],'4': [],'5': ['Start'],'6': [],'7': [],'8': []},
			checkedCheckboxes: '',
			progressBar: '',
			siteUrl: ''
		}, options);
		
		layout_load();
		
		function getWorkflowStatus(siteUrl, selectedStatus, theContainer) {
			var users = [];
			var context = new SP.ClientContext(siteUrl);
			var workflowServicesManager = new SP.WorkflowServices.WorkflowServicesManager(context, context.get_web());
			
			//Workflow Subscriptions
			var workflowSubscriptions = workflowServicesManager.getWorkflowSubscriptionService().enumerateSubscriptions();
			context.load(workflowSubscriptions);
			
			execute(context)
				.done(function() {
					var workflowCollection = [];
		
					for(var i = 0; i < workflowSubscriptions.get_count(); i++) {
						var workflowSubscription = workflowSubscriptions.get_item(i);
						var workflowInstance = workflowServicesManager.getWorkflowInstanceService().enumerate(workflowSubscription);
						context.load(workflowInstance);

						workflowCollection.push({ "instance" : workflowInstance, "subscription" : workflowSubscription });
					}
					
					execute(context)
						.done(function() {
							var listGUID = '';

							for(var i = 0; i < workflowCollection.length; i++) {
								listGUID = workflowCollection[i].subscription.get_propertyDefinitions()["Microsoft.SharePoint.ActivationProperties.ListId"];
								
								// Create DIV if it does not exist, add h4 and table
								if(!theContainer.children('div[id="' + listGUID + '-div"]').length) {
									theContainer.append(template({
										name: 'rb-listContainer', 
										divId: listGUID + '-div', 
										header: workflowCollection[i].subscription.get_propertyDefinitions()["Microsoft.SharePoint.ActivationProperties.ListName"], 
										tableId: listGUID + '-table'
									}));
								}
		
								var data = workflowCollection[i].instance.get_data();
								var history = [];
								for(var j = 0; j < data.length; j++) {
									var itemId = data[j].get_properties()["Microsoft.SharePoint.ActivationProperties.ItemId"];
									var r = new rbWorkflowStatus(data[j].get_status(), data[j].get_id());
										r.buttonCollection = settings.buttonCollection;
									var initiatorUserId = data[j].get_properties()["Microsoft.SharePoint.ActivationProperties.InitiatorUserId"];
									
									if(initiatorUserId != undefined) {
										if(!$.grep(users, function(e) { return e.id == initiatorUserId.split('\\')[1]; }).length) {
											var user = context.get_web().ensureUser(initiatorUserId.split('\\')[1]);
											context.load(user);
											
											users.push({ "id" : initiatorUserId.split('\\')[1], "user": user });
										}
									}

									history.push({ 
										"itemId": itemId,
										"initiatorUserId": initiatorUserId,
										"instanceCreated": getDate(new Date(data[j].get_instanceCreated())),
										"lastUpdated": getDate(new Date(data[j].get_lastUpdated())),
										"statusId": r.get_id(),
										"statusText": r.get_text()
									});

									var d = {};
										d.history = $.grep(history, function(e) { return e.itemId === itemId; });
										d.initiatorUserId = initiatorUserId;
										d.instanceGuid = data[j].get_id().toString();
										d.itemId = data[j].get_properties()["Microsoft.SharePoint.ActivationProperties.ItemId"];
										d.subscriptionGuid = data[j].get_workflowSubscriptionId().toString();

									// Add a new row to the table using the Workflow Instance data, 
									if(r.get_inArray(selectedStatus) != -1) {
										$('#' + listGUID + '-table tbody').append(template({
											name: 'rb-tableRow',
											initiatedOn: getDate(new Date(data[j].get_instanceCreated())),
											initiatorUserId: (initiatorUserId != undefined ? initiatorUserId.split('|')[1] : initiatorUserId),
											itemUrl: siteUrl + '/' + data[j].get_properties()["Microsoft.SharePoint.ActivationProperties.CurrentItemUrl"],
											itemId: itemId,
											workflowName: workflowCollection[i].subscription.get_name(),
											button: r.get_buttonHTML(d),
											internalStatus: r.get_text(),
											userStatus: data[j].get_userStatus(),
											faultInfo: data[j].get_faultInfo()
										}));
									}
								}
							}
							
							execute(context)
								.done(function() {
									for(var i = 0; i < users.length; i++) {										
										$(theContainer).find('td > div:contains("' + users[i].id + '")').replaceWith('Initiated By: ' + users[i].user.get_title() + ' (' + users[i].id + ')');
									}
								})
						})	
						.then(function() {
							$('.rb-progressBar').hide();
							$(theContainer).children('.rb-listContainer').show();
							
							//Bind click event to Resume button
							$('div[id*="button-Resume"]').click(function() {
								runWIS('resumeWorkflow', JSON.parse($(this).attr('data-value')), siteUrl, theContainer);
							});
							
							//Bind click event to Start button
							$('div[id*="button-Start"]').click(function() {
								var d = JSON.parse($(this).attr('data-value'));
								var h = $.grep(d.history, function(e) { return e.statusText === 'Completed'; });

								if(h.length) {
									SP.UI.ModalDialog.showModalDialog({
										title: 'Start the Workflow',
										html: dialogMessageHistory(h, true),
										dialogReturnValueCallback: function(result) {
											if(result === SP.UI.DialogResult.OK) {
												startWorkflow(JSON.parse($(this).attr('data-value')), siteUrl, theContainer);
											}
										}
									});
								} else {
									startWorkflow(JSON.parse($(this).attr('data-value')), siteUrl, theContainer);
								}
							});
							
							//Bind click event to Terminate button
							$('div[id*="button-Terminate"]').click(function() {
								var d = JSON.parse($(this).attr('data-value'));
								
								SP.UI.ModalDialog.showModalDialog({
									title: 'End the Workflow',
									html: dialogMessage('Do you want to end the workflow?', true),
									dialogReturnValueCallback: function(result) {
										if(result === SP.UI.DialogResult.OK) {
											runWIS('terminateWorkflow', d, siteUrl, theContainer);
										}
									}
								});
							});
		
						})
						.fail(function(error) {
							onError(error, siteUrl, theContainer);
						});		
				})
				.fail(function(error) {
					onError(error, siteUrl, theContainer);
				});
		}
		
		// =============================================================
		// runWIS
		// Run WorkflowInstanceService methods
		// Methods, the first parameter: 
		//		resumeWorkflow
		//		terminateWorkflow
		// =============================================================
		function runWIS(method, d, siteUrl, theContainer) {
			var dlg = SP.UI.ModalDialog.showWaitScreenWithNoClose("Please wait...", "Waiting for workflow...", null, null);
			var context = new SP.ClientContext(siteUrl);
			var workflowServicesManager = new SP.WorkflowServices.WorkflowServicesManager(context, context.get_web());
			
			//Workflow Instance
			var instance = workflowServicesManager.getWorkflowInstanceService().getInstance(d.instanceGuid);
		
			workflowServicesManager.getWorkflowInstanceService()[method](instance);
			execute(context)
				.done(function() {
					setTimeout(function() {
						dlg.close();
						$('#button-get-status').trigger('click');
					}, 2000);
				})
				.fail(function(error) {
					dlg.close();
					onError(error, siteUrl, theContainer);
				});
		}

		function startWorkflow(d, siteUrl, theContainer) {
			var dlg = SP.UI.ModalDialog.showWaitScreenWithNoClose("Please wait...", "Waiting for workflow...", null, null);
			var context = new SP.ClientContext(siteUrl);
			var workflowServicesManager = new SP.WorkflowServices.WorkflowServicesManager(context, context.get_web());
			var subscription = workflowServicesManager.getWorkflowSubscriptionService().getSubscription(d.subscriptionGuid);
			
			context.load(subscription, 'PropertyDefinitions');
			execute(context)
				.done(function() {
					var params = new Object();
					var formData = subscription.get_propertyDefinitions()["FormData"];
					
					if(formData != null && formData != 'undefined' && formData != "") {
						var assocParams = formData.split(";#");
						
						for(var i = 0; i < assocParams.length; i++) {
							params[assocParams[i]] = subscription.get_propertyDefinitions()[assocParams[i]];
						}
					}
					
					if(d.itemId) {
						workflowServicesManager.getWorkflowInstanceService().startWorkflowOnListItem(subscription, d.itemId, params);
					} else {
						workflowServicesManager.getWorkflowInstanceService().startWorkflow(subscription, params);
					}
					
					execute(context)
						.done(function() {
							setTimeout(function() {
								dlg.close();
								$('#button-get-status').trigger('click');
							}, 2000);
						})
						.fail(function(error) {
							dlg.close();
							onError(error, siteUrl, theContainer);
						});
				})
				.fail(function(error) {
					dlg.close();
					onError(error, siteUrl, theContainer);
				});
		}

		// =============================================================
		// execute
		// Use to perform executeQueryAsync with Deferred
		// =============================================================
		function execute(clientContext) {
			var deferred = $.Deferred();
			
			clientContext.executeQueryAsync(
				function(sender, args) { 
					deferred.resolve();
				},
				function(sender, args) {
					deferred.reject(args);
				}
			);
			
			return deferred;
		}
		
		function getDate(d) {
			d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
			return d.format('M/d/yyyy hh:mm:ss tt');
		}
		
		// =============================================================
		// onError
		// =============================================================
		function onError(error, siteUrl, theContainer) {
			var errorText = {
				'404': '404 (Not Found).'
			};
			var code = error.get_message().match(/\d+/) || '';
			var errorMessage = (code.length > 0 ? ' - ' + errorText[code] : '');
			$('.rb-progressBar').hide();
			theContainer.append(template({name: 'rb-error', headerMsg: errorMessage, errorMsg: siteUrl + errorMessage, detailMsg: error.get_message()}));
		}

		function layout_load() {
			// Add UL
			$(theContainer).append($('<ul class="rb-input-wrapper">'));
			
			// Add site Url input 
			$(theContainer).find('ul.rb-input-wrapper').append(
				$('<li class="rb-input-item">').append(
					$('<div class="rb-label">').text('Site Url:'),
					$('<div class="rb-input">').html('<input id="rb-input-siteUrl" value="' + settings.siteUrl + '">')
				)
			);
			
			// Add Workflow Status fieldset
			$(theContainer).find('ul.rb-input-wrapper').append(
				$('<li class="rb-input-item">').append(
					$('<fieldset>').append(
						$('<legend>').text('Choose Workflows Status')
					)
				)
			);
			// Add Workflow Status checkboxes
			var statusCollection = (new rbWorkflowStatus()).statusCollection;
			for(var i = 0; i < statusCollection.length; i++) {
				if(i % 3 == 0) {
					$(theContainer).find('li.rb-input-item fieldset').append($('<div class="rb-section">'));
				}
									
				$(theContainer).find('li.rb-input-item fieldset div.rb-section:last').append(
					$('<div>').append(
						$('<input type="checkbox" id="checkbox' + i + '" name="workflowStatus" value="' + i + '" ' + (settings.checkedCheckboxes.indexOf(statusCollection[i]) != -1 ? 'checked="checked"' : '') + '>'),
						$('<label for="' + i + '">').text(statusCollection[i])
					)
				);
			}
			
			// Add button
			$(theContainer).find('ul.rb-input-wrapper').append(
				$('<li class="rb-input-item">').append(
					$('<div class="rb-button-wrapper">').append(
						$('<div id="button-get-status" style="overflow: hidden; text-align: center;">').append(
							$('<h4 class="rb-button">').text('Get Status')
						)
					)
				)
			);

			// Make sure enter key is Canceled
			$('#rb-input-siteUrl').keydown(function(e) {
				if(e.keyCode == 13) {
					return false;
				}
			});
			
			// Add progressBar
			$(theContainer).append(
				$('<div class="rb-progressBar" style="display:none;">').append(settings.progressBar)
			);
		
			// Bind Click event to the button
			$('#button-get-status').click(function() {
				theContainer.children('.rb-listContainer').remove();
				$('.rb-progressBar').show();
				
				var siteUrl = $('#rb-input-siteUrl').val();
				var selectedStatus = [];
				
				if(siteUrl.length) {
					$('input[name="workflowStatus"]:checked').each(function() {
						selectedStatus.push(parseInt($(this).val()));
					});
					
					getWorkflowStatus(siteUrl, selectedStatus, theContainer);
				} else {
					alert('Site Url is blank. Please enter a valid Site Url and try again');
				}
			});
		}
		
		// =============================================================
		// template
		//		rb-button
		//		rb-error
		//		rb-listContainer
		//		rb-tableRow
		// =============================================================
		function template(obj) {
			var r = '', i, u;
			
			templates = {
				"rb-button" : function(obj) {
					with(obj) {
						r = '<div class="' + ((i = cssClassWrapper) == null ? 'rb-button-wrapper-small' : i) + '">';
						r +='	<div id="button-' + ((i = buttonId) == null ? 'undefined' : i) + '" data-value="' + ((i = dataValue) == null ? '' : i) + '"  style="overflow: hidden; text-align: center;">';
						r += '		<h4 class="rb-button ' + ((i = cssClass) == null ? '' : i) + '">' + ((i = buttonText) == null ? 'undefined' : i) + '</h4>';
						r += '	</div>';
						r += '</div>';
					}
					return r;
				},
				"rb-error" : function(obj) {
					with(obj) {
						r = '<div class="rb-listContainer">';
						r += '	<h2>Error' + ((i = headerMsg) == null ? '' : i) + '</h2>';
						r += '	<h3>' + ((i = errorMsg) == null ? '' : i) + '</h3>';
						r += '	<div">Details: ' + ((i = detailMsg) == null ? '' : i) + '</div>';
						r += '</div>';
					}
					return r;
				},
				"rb-listContainer" : function(obj) {
					with(obj) {
						r = '<div id="' + ((i = divId) == null ? '' : i) + '" class="rb-listContainer" style="display:none;">';
						r += '	<h2>' + ((i = header) == null ? '' : i) + '</h2>';
						r += '	<table id="' + ((i = tableId) == null ? '' : i) + '">';
						r += '		<thead>';
						r += '			<th>Item ID</th>';
						r += '			<th>Workflow</th>';
						r += '			<th>Internal Status</th>';
						r += '			<th>Status</th>';
						r += '			<th>Notes</th>';
						r += '		</thead>';
						r += '		<tbody></tbody>';
						r += '	</table>';
						r += '</div>';
					}
					return r;
				},
				"rb-tableRow" : function(obj) {
					with(obj) {
						r = '<tr>';
						r += '	<td><a href="' + ((i = itemUrl) == null ? '' : i) + '">' + ((i = itemId) == null ? '' : i) + '</a></td>';
						r += '	<td><div>' + ((i = workflowName) == null ? '' : i) + '</div>' + ((i = button) == null? '' : i) + '</td>';
						r += '	<td>' + ((i = internalStatus) == null ? '' : i) + '</td>';
						r += '	<td>' + ((i = userStatus) == null ? '' : i) + '</td>';
						r += '	<td><div>Initiated On: ' + initiatedOn + '</div><div>Initiated By: ' + initiatorUserId + '</div>' + ((i = faultInfo) == null ? '' : '<div>Fault Info:<br>' + i + '</div>') + '</td>';
						r += '</tr>';
					}
					return r;
				}
			};
			return templates[obj.name](obj);
		}
	
		// =============================================================
		// Dialog Screen
		//		btnCancel_onClick
		//		btnOK_onClick
		//		dialogMessage
		// =============================================================
		function btnCancel_onClick() {
			SP.UI.ModalDialog.commonModalDialogClose(SP.UI.DialogResult.Cancel);
		}
		
		function btnOK_onClick() {
			SP.UI.ModalDialog.commonModalDialogClose(SP.UI.DialogResult.OK);
		}
		
		function dialogMessage(msgs, confirm) {
			var el = document.createElement('div');
			
			el.style.cssText = 'display:block; width:100%';
			if($.type(msgs) === 'array') {
				el.innerHTML = '<ul>';
				$.each(msgs, function(idx, val) {
					el.innerHTML += '<li>' + val + '</li>';
				});
				el.innerHTML += '</ul>';
		
			} else if($.type(msgs) === 'string') {
				el.innerHTML += msgs;
			}
		
			//If Confirm, add the Cancel button. Otherwise, just OK button.
			var btnHtml = '<div style="width:100%; text-align:right">';
				btnHtml += '<input type="button" value="OK" onclick="javascript:SP.UI.ModalDialog.commonModalDialogClose(SP.UI.DialogResult.OK);" />';
				btnHtml += (confirm ? '<input type="button" value="Cancel" onClick="javascript:SP.UI.ModalDialog.commonModalDialogClose(SP.UI.DialogResult.Cancel);" />' : '') + '</div>';
			el.innerHTML += btnHtml;
			return el;
		}
		
		function dialogMessageHistory(h, confirm) {
			var m = 'This workflow has been instantiated with the following details:<br/>';
			for(var i = 0; i < h.length; i++) {
				m += '<ul><li>Instance Created on: ' + h[i].instanceCreated;
				m += '<ul><li>Last Updated on: ' + h[i].lastUpdated + '</li>';
				m += '<li>Status: ' + h[i].statusText + '</li></ul></ul>';
			}
			m += '<div style="padding-top:16px;padding-bottom:16px;">Are you sure you want to start the workflow?</div>';
			
			return dialogMessage(m, confirm);
		}
		/* end of Dialog Screen */
	}
	
	// =============================================================
	// rbWorkflowStatus class
	// Properties: id, instanceGUID, statusCollection, buttonCollection
	// Methods: get_buttonHTML, get_id, get_inArray, get_text()
	// =============================================================
	class rbWorkflowStatus {
		constructor(id, instanceGUID) {
			this.id = id;
			this.instanceGUID = instanceGUID;
		
			this.statusCollection = ['Not Started','Started','Suspended','Canceling','Canceled','Terminated','Completed','Not Specified','Invalid'];
			this.buttonCollection = {'0': [],'1': [],'2': ['Resume', 'Terminate'],'3': [],'4': [],'5': ['Start'],'6': [],'7': [],'8': []};
		}
		
		get_buttonHTML(d) {
			function template(text, id, dataValue) {
				var r = '<div class="rb-button-wrapper-small">';
					r +='	<div id="button-' + text + '-' + id + '" data-value=\'' + JSON.stringify(dataValue) + '\'  style="overflow: hidden; text-align: center;">';
					r += '		<h4 class="rb-button rb-button-' + text.toString().toLowerCase() + '">' + text + '</h4>';
					r += '	</div>';
					r += '</div>';
				return r;
			}
			
			var u = '';
			
			if(d.itemId > 0 && d.itemId != undefined) {
				var buttons = this.buttonCollection[this.id];
				for(var i = 0; i < buttons.length; i++) {
					u += template(buttons[i], this.instanceGUID, d);
				}
			} else if(this.id == 2) {
				u += template('Terminate', this.instanceGUID, d);
			}
			
			return u;
		}
		
		get_id() {
			return this.id;
		}
	
		get_inArray(array, allowBlank) {
			if(!array.length && (true || allowBlank)) {
				return this.id;
			} else if(array.length && Array.isArray(array)) {
				return array.indexOf(this.id);
			} else {
				return -1;
			}
		}
		
		get_text() {
			var r = parseInt(this.id);
		
			return (r != NaN && r != null && r != undefined && r < 8 ? this.statusCollection[r] : 'Unknown');
		}
	}
	// end of rbWorkflowStatus class
});