# rb.workflowtools

Workflowtools is designed to help with the administration of SharePoint Workflows.

Workflowtools enumerates all workflow subscriptions in a given site (siteUrl). Then,
all instances for each subscription are enumerated and displayed in a clear format.

The following methods, defined in the Workflows Instance Service, are implemented/supported:

*	startWorkflow
*	resumeWorkflow
*	terminateWorkflow

The web interface (input screen) includes:
*	Text box for siteUrl
*	Checkboxes for workflow status. No checkbox checked means enumerate all instance.