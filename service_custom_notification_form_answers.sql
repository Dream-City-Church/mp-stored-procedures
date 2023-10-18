USE [MinistryPlatform]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[service_custom_notification_form_answers]

	@DomainID INT

AS

/****************************************************
*** Form Answers Email Notification ***
*****************************************************
A custom Dream City Church procedure for Ministry Platform
Version: 1.1
Author: Stephan Swinford
Date: 03/25/2022

This procedure is provided "as is" with no warranties expressed or implied.

-- Description --

Sends an email to the Form Primary Contact with the form answers.
Add this procedure to a 15 minute SQL Server Agent Job.

You may need to disable the standard form notification job
to avoid duplicate notifications.


*****************************************************
****************** BEGIN PROCEDURE ******************
*****************************************************/

-- Start with setting our procedure variables
DECLARE
-- These variables are useful for testing
    @TestMode BIT = 0 -- 0 is regular operation, 1 will run the procedure in test mode without sending any email (console output only)
    ,@TestEmail BIT = 0 -- 0 is regular operation, 1 will send all emails to @TestEmailAddress instead of the Event Primary Contact
    ,@TestEmailAddress VARCHAR(100) = 'sswinford@dreamcitychurch.us' -- Email address that you want to receive emails when @TestEmail is set to 1

-- These variables are set from Configuration Setting Keys
    ,@MessageID INT = (SELECT top 1 Value FROM dp_Configuration_Settings CS WHERE ISNUMERIC(Value) = 1 AND CS.Domain_ID = @DomainID AND CS.Application_Code = 'Services' AND Key_Name = 'NotificationFormResponseAnswersMessageID')

-- And these variables are used later in the procedure
	,@ContactID INT = 0
    ,@EmailTo VARCHAR (500)
    ,@EmailSubject VARCHAR(1000)
    ,@EmailBody VARCHAR(MAX)
    ,@FormAnswersList VARCHAR(MAX) = ''
    ,@EmailFrom VARCHAR(500)
    ,@EmailReplyTo VARCHAR(500)
    ,@FormResponseID INT = 0
	,@FormTitle VARCHAR(500)
	,@ResponseFirstName VARCHAR(500)
	,@ResponseLastName VARCHAR(500)
	,@ResponseDate DateTime
	,@ResponsePhone VARCHAR(500)
	,@ResponseEmail VARCHAR(500)
	,@ResponseComments VARCHAR(500)
    ,@CopyMessageID INT
    ,@BaseURL NVARCHAR(250) = ISNULL((SELECT Top 1 Value FROM dp_Configuration_Settings CS WHERE CS.Domain_ID = @DomainID AND CS.Application_Code = 'SSRS' AND CS.Key_Name = 'BASEURL'),'')
	,@ExternalURL NVARCHAR(250) = ISNULL((SELECT Top 1 External_Server_Name FROM dp_Domains WHERE dp_Domains.Domain_ID = @DomainID),'')
	,@FormResponsePageID INT = ISNULL((SELECT TOP 1 Page_ID FROM dp_Pages P WHERE P.Table_Name = 'Form_Responses' AND P.Filter_Clause IS NULL ORDER BY Page_ID),'')
    
-- Check that the template Message ID actually exists and our key values are not NULL before running the procedure
IF EXISTS (
		SELECT 1 
		FROM dp_Communications Com 
		WHERE Com.Communication_ID = @MessageID 
		 AND Com.Domain_ID = @DomainID
		 )
    AND EXISTS (
		SELECT TOP(1) *
		FROM Form_Responses FR 
		 LEFT JOIN Forms F ON FR.Form_ID = F.Form_ID 
		WHERE F.Notify = 1 
		 AND FR._Notification_Sent_Date IS NULL 
		 AND FR.Response_Date >= GetDate()-7
		 )
BEGIN

-- Create some temp tables for audit log
	CREATE TABLE #CommAudit (Communication_ID INT)
	CREATE TABLE #MessageAudit (Communication_Message_ID INT)

-- Set some initial variables based on the template
	SET @EmailBody = ISNULL((SELECT Top 1 Body FROM dp_Communications C WHERE C.Communication_ID = @MessageID),'')
	SET @EmailSubject = ISNULL((SELECT Top 1 Subject FROM dp_Communications C WHERE C.Communication_ID = @MessageID),'')
    SET @EmailFrom = ISNULL((SELECT Top 1 '"' + Nickname + ' ' + Last_Name + '" <' + Email_Address + '>' FROM Contacts C LEFT JOIN dp_Communications Com ON Com.From_Contact = C.Contact_ID WHERE C.Contact_ID = Com.From_Contact AND Com.Communication_ID = @MessageID),'')
    SET @EmailReplyTo = ISNULL((SELECT Top 1 '"' + Nickname + ' ' + Last_Name + '" <' + Email_Address + '>' FROM Contacts C LEFT JOIN dp_Communications Com ON Com.Reply_to_Contact = C.Contact_ID  WHERE C.Contact_ID = Com.Reply_to_Contact AND Com.Communication_ID = @MessageID),'')

-- Create our Message record. We'll add recipients later.
	INSERT INTO [dbo].[dp_Communications]
		([Author_User_ID]
		,[Subject]
		,[Body]
		,[Domain_ID]
		,[Start_Date]
		,[Expire_Date]
		,[Communication_Status_ID]
		,[From_Contact]
		,[Reply_to_Contact]
		,[_Sent_From_Task]
		,[Selection_ID]
		,[Template]
		,[Active]
		,[To_Contact]) 
	OUTPUT INSERTED.Communication_ID
	INTO #CommAudit
	SELECT [Author_User_ID]
		,@EMailSubject AS [Subject]
        ,@EmailBody AS Body
		,[Domain_ID]
		,[Start_Date] = GETDATE() 
		,[Expire_Date]
		,[Communication_Status_ID] = 3
		,[From_Contact]
		,[Reply_to_Contact]
		,[_Sent_From_Task] = NULL 
		,[Selection_ID] = NULL
		,[Template] = 0
		,[Active] = 0
		,[To_Contact]
	FROM  dp_Communications Com 
	WHERE Com.Communication_ID = @MessageID 

-- Set this variable based on the Message record we just created above
    SET @CopyMessageID = SCOPE_IDENTITY()

-- Insert into the Audit Log table --
	INSERT INTO dp_Audit_Log (Table_Name,Record_ID,Audit_Description,User_Name,User_ID,Date_Time)
	SELECT 'dp_Communications',#CommAudit.Communication_ID,'Created','Svc Mngr',0,GETDATE() 
	FROM #CommAudit

-- Create our cursor list (recipient list)
    DECLARE CursorEmailList CURSOR FAST_FORWARD FOR
	    SELECT Contact_ID = C.Contact_ID
	        ,Email_To = ISNULL('"' +  C.Nickname + ' ' + C.Last_Name + '" <' + C.Email_Address + '>','')
	        ,Email_Subject = REPLACE(@EmailSubject,'[Form_Title]', F.Form_Title)
	        ,Email_Body = REPLACE(REPLACE(REPLACE(@EmailBody,'[Nickname]',ISNULL(C.Nickname,C.Display_Name)),'[BaseURL]',@BaseURL),'[Response_URL]','https://' + @ExternalURL + '/mp/' + CAST(@FormResponsePageID AS VARCHAR) + '/' + CAST(FR.Form_Response_ID AS VARCHAR))
            ,FormResponseID = FR.Form_Response_ID
			,ResponseFirstName = ISNULL(FR.First_Name,C2.Nickname)
			,ResponseLastName = ISNULL(FR.Last_Name,C2.Last_Name)
			,ResponseDate = FR.Response_Date
			,FormTitle = F.Form_Title
			,ResponsePhone = ISNULL(FR.Phone_Number,C2.Mobile_Phone)
			,ResponseEmail = ISNULL(FR.Email_Address,C2.Email_Address)
            ,ResponseComments = R.Comments
	    FROM Form_Responses FR
			LEFT JOIN Forms F ON FR.Form_ID=F.Form_ID
			LEFT JOIN Contacts C ON F.Primary_Contact=C.Contact_ID
			LEFT JOIN Contacts C2 ON FR.Contact_ID=C2.Contact_ID
			LEFT JOIN Opportunities O ON O.Custom_Form=F.Form_ID
			LEFT JOIN Responses R ON R.Opportunity_ID=O.Opportunity_ID
			LEFT JOIN Participants P ON R.Participant_ID=P.Participant_ID
        WHERE C.Email_Address IS NOT NULL
			AND ((ISNULL(P.Participant_ID,C2.Participant_Record) = C2.Participant_Record) OR (ISNULL(P.Participant_ID,C2.Participant_Record) IS NULL))
			AND ISNULL(R.Response_Date,GetDate()) >= GetDate()-30
            AND F.Notify = 1
			AND FR.Response_Date >= GetDate()-7
			AND FR._Notification_Sent_Date IS NULL
	        AND C.Domain_ID = @DomainID

-- Now lets open the cursor list and create notifications from it
    OPEN CursorEmailList
	FETCH NEXT FROM CursorEmailList INTO @ContactID, @EmailTo, @EmailSubject, @EmailBody, @FormResponseID, @ResponseFirstName, @ResponseLastName, @ResponseDate, @FormTitle, @ResponsePhone, @ResponseEmail, @ResponseComments
		WHILE @@FETCH_STATUS = 0
			BEGIN
            -- We initially set the @FormAnswersList variable with some opening HTML for the email template
                SET @FormAnswersList = '<ol>'

                IF @ResponseComments IS NOT NULL BEGIN
                   SET @FormAnswersList = COALESCE(@FormAnswersList + '<span style="padding:5px 5px 15px 2px;font-weight:bold;">Comments: </span>' + @ResponseComments,'')
                END
                
            -- And then concatenate onto that the details for each individual event. You can modify the HTML to adjust the styling of the table
                SELECT @FormAnswersList = COALESCE(@FormAnswersList + '<li style="margin-bottom:1em;"><span style="font-weight:bold;">' + FF.Field_Label + ':</span><br />' + 
						CASE WHEN FRA.Response = 'File(s) attached to record'
							THEN '<a href="https://' + @ExternalURL + '/ministryplatformapi/files/'+(SELECT TOP(1) CAST(FIL.Unique_Name AS VARCHAR(50)) FROM dp_Files FIL WHERE FIL.Table_Name = 'Form_Response_Answers' AND FIL.Record_ID = FRA.Form_Response_Answer_ID)+'" >Download File Here</a>'
							ELSE ISNULL(FRA.Response,'')
						 END 
						 + '</li>','')
				    FROM Form_Response_Answers FRA
                        LEFT JOIN Form_Fields FF ON FF.Form_Field_ID = FRA.Form_Field_ID 
                        LEFT JOIN Forms F ON FF.Form_ID = F.Form_ID
                        LEFT JOIN Form_Responses FR ON FRA.Form_Response_ID = FR.Form_Response_ID
				    WHERE FR.Form_Response_ID = @FormResponseID
						AND FR._Notification_Sent_Date IS NULL
					ORDER BY FF.Field_Order

                SET @FormAnswersList = @FormAnswersList + '</ol>'

			-- Check our @TestMode flag if 0 for normal operation
                IF @TestMode = 0
	                BEGIN
                        
                    -- Replace placeholder in email template with our content
                        SET @EmailBody = ISNULL(REPLACE(@EmailBody,'[Form_Response_Answers]',@FormAnswersList),@EmailBody)
						SET @EmailBody = ISNULL(REPLACE(@EmailBody,'[Web_First_Name]',@ResponseFirstName),@EmailBody)
						SET @EmailBody = ISNULL(REPLACE(@EmailBody,'[Web_Last_Name]',@ResponseLastName),@EmailBody)
						SET @EmailBody = ISNULL(REPLACE(@EmailBody,'[Response_Date]',@ResponseDate),@EmailBody)
						SET @EmailBody = ISNULL(REPLACE(@EmailBody,'[Form_Title]',@FormTitle),@EmailBody)
						SET @EmailBody = ISNULL(REPLACE(@EmailBody,'[Web_Phone_Number]',ISNULL(@ResponsePhone,'')),@EmailBody)
						SET @EmailBody = ISNULL(REPLACE(@EmailBody,'[Web_Email_Address]',ISNULL(@ResponseEmail,'')),@EmailBody)
						SET @EmailBody = ISNULL(REPLACE(@EmailBody,'[Email_Address]',ISNULL((SELECT Email_Address FROM Contacts C WHERE C.Contact_ID = @ContactID),'')),@EmailBody)
						

                    -- We'll now add Recipients to the Message we created earlier
                        INSERT INTO [dbo].[dp_Communication_Messages]
				            ([Communication_ID]
				            ,[Action_Status_ID]
				            ,[Action_Status_Time]
				            ,[Action_Text]
				            ,[Contact_ID]
				            ,[From]
				            ,[To]
				            ,[Reply_To]
				            ,[Subject]
				            ,[Body]
				            ,[Domain_ID]
				            ,[Deleted])
						OUTPUT INSERTED.Communication_Message_ID
						INTO #MessageAudit
		                SELECT DISTINCT [Communication_ID] = @CopyMessageID 
				            ,[Action_Status_ID] = 2
				            ,[Action_Status_Time] = GETDATE()
				            ,[Action_Text] = NULL
				            ,[Contact_ID] = @ContactID 
				            ,[From] = @EmailFrom
				            ,[To] = CASE WHEN @TestEmail = 0 THEN @EmailTo ELSE @TestEmailAddress END
				            ,[Reply_To] = @ResponseEmail
				            ,[Subject] = @EmailSubject
				            ,[Body] = @EmailBody
				            ,[Domain_ID] = @DomainID
				            ,[Deleted] = 0

						UPDATE Form_Responses
							SET _Notification_Sent_Date = GetDate()
							WHERE Form_Response_ID = @FormResponseID

		            END
				ELSE
                -- If we're testing, just print out these results in console
					BEGIN
						SELECT @ContactID, @EmailTo, @EmailSubject, @EmailBody, @FormResponseID, @ResponseFirstName, @ResponseLastName, @ResponseDate, @FormTitle, @ResponsePhone, @ResponseEmail, @ResponseComments
					END

        -- And now move on to the next Recipient in our cursor list
			FETCH NEXT FROM CursorEmailList INTO  @ContactID, @EmailTo, @EmailSubject, @EmailBody, @FormResponseID, @ResponseFirstName, @ResponseLastName, @ResponseDate, @FormTitle, @ResponsePhone, @ResponseEmail, @ResponseComments

    -- Done fetching from the cursor list
        END

-- Done with the cursor list, so let's make sure it's cleared
	DEALLOCATE CursorEmailList

-- Create an audit log entry for each Communication Message that was created
	INSERT INTO dp_Audit_Log (Table_Name,Record_ID,Audit_Description,User_Name,User_ID,Date_Time)
	SELECT 'dp_Communication_Messages',#MessageAudit.Communication_Message_ID,'Created','Svc Mngr',0,GETDATE() 
	FROM #MessageAudit

-- Drop our temporary tables that we were using for audit logging
	DROP TABLE #CommAudit
	DROP TABLE #MessageAudit

-- Done with our initial 'if template exists'
END

