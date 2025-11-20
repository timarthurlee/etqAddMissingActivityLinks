def addMissingActivityLinks(subformName, activityFormName, fieldMap, document=None, debugEnabled=False):
	'''
	Adds missing Activity Links to records in the specified subform of the given document.
	
	subformName: Name of the subform containing the records to check.
	activityFormName: Name of the Activity form to link to.
	fieldMap: Dictionary mapping subform field names to Activity form database column names for matching records to corresponding Activity documnents [activity subform field name -> activity document DB column name].
	document: Document object to process. If None, uses thisDocument.
	debugEnabled: Boolean to enable or disable debug logging.
	'''

	document = thisDocument if document is None else document
	output = False

	fDebug = EtqDebug('addMissingActivityLinks()', document=document, enabled=debugEnabled)
	fDebug.log([subformName, activityFormName, fieldMap, document], label='Function Parameters')

	fieldMapItems = list(fieldMap.items())

	application = document.getParentApplication()
	schema = application.getSchemaName()
	
	# Form
	formSetting = document.getFormSetting()	
	formName = formSetting.getName()
	
	# Activity form
	activityFormSetting = PublicSettingManager().getFormSetting(activityFormName)
	activityFormTable = activityFormSetting.getTableName()
	activityFormKey = activityFormSetting.getPrimaryKey()
	ACTIVITY_LINK_FIELD_NAME = 'ETQ$SUBFORM_ACTIVITY_LINK'

	def norm(v):
		return '' if v in (None, 'None', 'null') else v
	
	# Checking for missing and existing Activity Links
	activitySubform = document.getSubform(subformName)
	missing = {}
	foundActivities = []
	for rec in activitySubform.getRecords():
		recID = rec.getRecordID()
		activityLinks = rec.getField(ACTIVITY_LINK_FIELD_NAME).getDocLinks()
		if activityLinks:	
			# Activity link exists, recording Action Item ID	
			activityID  = str(activityLinks[0].getDocKey().getKeyValue())
			foundActivities.append(activityID)
			fDebug.log('Activity Link already created for this record, storing... [Rec ID: {recID} -> Activity ID: {activityID}]'.format(recID=recID, activityID=activityID), label='Activity Rec Update Not Required')
		else:
			# No activity link found, adding to list for processing
			key = tuple(norm(rec.getFieldValue(k)) for k, v in fieldMapItems)
			missing.setdefault(key, rec)
			fDebug.log('No Activity Link found for this record, adding to processing list... [Rec ID: {}]'.format(recID), label='Activity Rec Update Required')

	if not missing:
		# No missing Activity Links found, exiting function
		fDebug.log('No missing Activity Links found, exiting function.', label='No Action Needed')
		return output
	
	# Building query to find matching Activity records
	exclude = ''
	if foundActivities:
		exclude = ' AND AF.{keyColumn} NOT IN ({found})'.format(keyColumn=activityFormKey, found=','.join(foundActivities))
	query = '\n'.join([
		"SELECT AF.{keyColumn}, AF.{columns} FROM {schema}.{tableName} AF",
		"LEFT JOIN {schema}.ETQ${tableName}_SL AF_SL ON AF.{keyColumn} = AF_SL.{keyColumn}",
		"LEFT JOIN {schema}.ETQ$DOCUMENT_LINKS DL ON DL.LINK_ID = AF_SL.ETQ$SOURCE_LINK",
		"LEFT JOIN ENGINE.FORM_SETTINGS SL_FORM ON DL.FORM_ID = SL_FORM.FORM_ID",
		"WHERE SL_FORM.FORM_NAME = '{parentFormName}' AND DL.DOCUMENT_ID = {parentDocumentID}{exclude}"
	]).format(schema=schema, tableName=activityFormTable, parentFormName=formName, parentDocumentID=document.getID(), columns=', AF.'.join(fieldMap.values()), keyColumn=activityFormKey, exclude=exclude)
	fDebug.log(query, label='Activity Query')
	dao = application.executeQueryFromDatasource("FILTER_ONLY", {"VAR$FILTER": query})
	# Finding matching Activity records
	while dao.next():					
		activityID = dao.getValue(activityFormKey)
		key = tuple(norm(dao.getValue(v)) for k, v in fieldMapItems)

		# Checking if Activity Document is not missing from Activity Subform
		if key not in missing:
			# No matching Activity record found, skipping.This can happen if Activities were removed from the Parent document activity subform.
			fDebug.log('No matching Activity record found for Activity ID {}, skipping.'.format(activityID), label='Skipping Activity Rec Update')
			continue
		
		rec = missing.get(key, None)

		# Checking if record is valid
		if not rec:
			# No Activity record found, skipping. This should not happen.
			fDebug.log('No Activity record found for Activity ID {}, skipping.'.format(activityID), label='Skipping Activity Rec Update')
			continue
		
		# Adding Activity Link to record
		recID = rec.getRecordID()

		fDebug.log('Document match found for Activity Rec. Rec ID: {}'.format(str(recID)) , label='Processing Activity Rec Update')
		activityLinks = application.getDocumentLinksByQuery(activityFormName, 'SELECT {keyColumn} FROM {schema}.{tableName} WHERE {keyColumn} = {documentID}'.format(schema=schema, tableName=activityFormTable, keyColumn=activityFormKey, documentID=activityID))
		if activityLinks:
			# Match found, adding Activity Link
			activityLink = activityLinks[0]
			activityLinkField = rec.getField(ACTIVITY_LINK_FIELD_NAME)
			fDebug.log('Adding Activity Link to Activity Rec. Rec ID: {} -> Doc ID: {}'.format(str(recID), str(activityID)), label='Adding Activity Link')
			activityLinkField.addDocLink(activityLink)
			if not output:
				output = True
		else:
			fDebug.log('No Document Link found for Activity ID {}, skipping.'.format(activityID), label='Skipping Activity Rec Update')
			continue
	return output