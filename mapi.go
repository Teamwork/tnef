package tnef

import "fmt"

// MAPIAttribute contains MAPI format attributes, i.e encoding type
// headers, attachments etc. See the constants for
// code references to find specific attributes.
type MAPIAttribute struct {
	Type int
	Name int
	Data []byte
	GUID int
}

func decodeMapi(data []byte) ([]MAPIAttribute, error) {
	var attrs []MAPIAttribute
	dataLen := len(data)
	offset := 0
	numProperties := byteToInt(data[offset : offset+4])
	offset += 4

	for i := 0; i < numProperties; i++ {
		if offset >= dataLen {
			continue
		}

		attrType := byteToInt(data[offset : offset+2])
		offset += 2

		isMultiValue := (attrType & mvFlag) != 0
		attrType &= ^mvFlag // Remove mvFlag

		typeSize := getTypeSize(attrType)
		if typeSize < 0 {
			isMultiValue = true
		}

		attrName := byteToInt(data[offset : offset+2])
		offset += 2

		guid := 0
		if attrName >= 0x8000 && attrName <= 0xFFFE {
			guid = byteToInt(data[offset : offset+16])
			offset += 16
			kind := byteToInt(data[offset : offset+4])
			offset += 4

			if kind == 0 {
				offset += 4
			} else if kind == 1 {
				iidLen := byteToInt(data[offset : offset+4])
				offset += 4

				offset += iidLen

				offset += (-iidLen & 3)
			}
		}

		// Handle multi-value properties
		valueCount := 1
		if isMultiValue {
			valueCount = byteToInt(data[offset : offset+4])
			offset += 4
		}

		if valueCount > 1024 && valueCount > len(data) {
			return nil, fmt.Errorf("count is too large: %d", valueCount)
		}

		attrData := []byte{}

		for i := 0; i < valueCount; i++ {
			length := typeSize
			if typeSize < 0 {
				length = byteToInt(data[offset : offset+4])
				offset += 4
			}

			// Read the data in
			attrData = append(attrData, data[offset:offset+length]...)

			offset += length
			offset += (-length & 3)
		}

		attrs = append(attrs, MAPIAttribute{Type: attrType, Name: attrName, Data: attrData, GUID: guid})
	}

	return attrs, nil
}

func getTypeSize(attrType int) int {
	switch attrType {
	case szmapiShort, szmapiBoolean:
		return 2
	case szmapiInt, szmapiFloat, szmapiError:
		return 4
	case szmapiDouble, szmapiApptime, szmapiCurrency, szmapiInt8byte, szmapiSystime:
		return 8
	case szmapiCLSID:
		return 16
	case szmapiString, szmapiUnicodeString, szmapiObject, szmapiBinary:
		return -1
	}
	return 0
}

const (
	mvFlag = 0x1000 // OR with type means multiple values

	szmapiUnspecified   = 0x0000 //# MAPI Unspecified
	szmapiNull          = 0x0001 //# MAPI null property
	szmapiShort         = 0x0002 //# MAPI short (signed 16 bits)
	szmapiInt           = 0x0003 //# MAPI integer (signed 32 bits)
	szmapiFloat         = 0x0004 //# MAPI float (4 bytes)
	szmapiDouble        = 0x0005 //# MAPI double
	szmapiCurrency      = 0x0006 //# MAPI currency (64 bits)
	szmapiApptime       = 0x0007 //# MAPI application time
	szmapiError         = 0x000a //# MAPI error (32 bits)
	szmapiBoolean       = 0x000b //# MAPI boolean (16 bits)
	szmapiObject        = 0x000d //# MAPI embedded object
	szmapiInt8byte      = 0x0014 //# MAPI 8 byte signed int
	szmapiString        = 0x001e //# MAPI string
	szmapiUnicodeString = 0x001f //# MAPI unicode-string (null terminated)
	szmapiPtSystime     = 0x001e //# MAPI time (after 2038/01/17 22:14:07 or before 1970/01/01 00:00:00)
	szmapiSystime       = 0x0040 //# MAPI time (64 bits)
	szmapiCLSID         = 0x0048 //# MAPI OLE GUID
	szmapiBinary        = 0x0102 //# MAPI binary
	szmapiUnknown       = 0x0033
)

// We can use these constants to find specific types
// of MAPIAttribute by comparing it to the type of the
// attribute.
const (
	MAPIAcknowledgementMode                   = 0x0001
	MAPIAlternateRecipientAllowed             = 0x0002
	MAPIAuthorizingUsers                      = 0x0003
	MAPIAutoForwardComment                    = 0x0004
	MAPIAutoForwarded                         = 0x0005
	MAPIContentConfidentialityAlgorithmID     = 0x0006
	MAPIContentCorrelator                     = 0x0007
	MAPIContentIdentifier                     = 0x0008
	MAPIContentLength                         = 0x0009
	MAPIContentReturnRequested                = 0x000A
	MAPIConversationKey                       = 0x000B
	MAPIConversionEits                        = 0x000C
	MAPIConversionWithLossProhibited          = 0x000D
	MAPIConvertedEits                         = 0x000E
	MAPIDeferredDeliveryTime                  = 0x000F
	MAPIDeliverTime                           = 0x0010
	MAPIDiscardReason                         = 0x0011
	MAPIDisclosureOfRecipients                = 0x0012
	MAPIDlExpansionHistory                    = 0x0013
	MAPIDlExpansionProhibited                 = 0x0014
	MAPIExpiryTime                            = 0x0015
	MAPIImplicitConversionProhibited          = 0x0016
	MAPIImportance                            = 0x0017
	MAPIIpmID                                 = 0x0018
	MAPILatestDeliveryTime                    = 0x0019
	MAPIMessageClass                          = 0x001A
	MAPIMessageDeliveryID                     = 0x001B
	MAPIMessageSecurityLabel                  = 0x001E
	MAPIObsoletedIpms                         = 0x001F
	MAPIOriginallyIntendedRecipientName       = 0x0020
	MAPIOriginalEits                          = 0x0021
	MAPIOriginatorCertificate                 = 0x0022
	MAPIOriginatorDeliveryReportRequested     = 0x0023
	MAPIOriginatorReturnAddress               = 0x0024
	MAPIParentKey                             = 0x0025
	MAPIPriority                              = 0x0026
	MAPIOriginCheck                           = 0x0027
	MAPIProofOfSubmissionRequested            = 0x0028
	MAPIReadReceiptRequested                  = 0x0029
	MAPIReceiptTime                           = 0x002A
	MAPIRecipientReassignmentProhibited       = 0x002B
	MAPIRedirectionHistory                    = 0x002C
	MAPIRelatedIpms                           = 0x002D
	MAPIOriginalSensitivity                   = 0x002E
	MAPILanguages                             = 0x002F
	MAPIReplyTime                             = 0x0030
	MAPIReportTag                             = 0x0031
	MAPIReportTime                            = 0x0032
	MAPIReturnedIpm                           = 0x0033
	MAPISecurity                              = 0x0034
	MAPIIncompleteCopy                        = 0x0035
	MAPISensitivity                           = 0x0036
	MAPISubject                               = 0x0037
	MAPISubjectIpm                            = 0x0038
	MAPIClientSubmitTime                      = 0x0039
	MAPIReportName                            = 0x003A
	MAPISentRepresentingSearchKey             = 0x003B
	MAPIX400ContentType                       = 0x003C
	MAPISubjectPrefix                         = 0x003D
	MAPINonReceiptReason                      = 0x003E
	MAPIReceivedByEntryID                     = 0x003F
	MAPIReceivedByName                        = 0x0040
	MAPISentRepresentingEntryID               = 0x0041
	MAPISentRepresentingName                  = 0x0042
	MAPIRcvdRepresentingEntryID               = 0x0043
	MAPIRcvdRepresentingName                  = 0x0044
	MAPIReportEntryID                         = 0x0045
	MAPIReadReceiptEntryID                    = 0x0046
	MAPIMessageSubmissionID                   = 0x0047
	MAPIProviderSubmitTime                    = 0x0048
	MAPIOriginalSubject                       = 0x0049
	MAPIDiscVal                               = 0x004A
	MAPIOrigMessageClass                      = 0x004B
	MAPIOriginalAuthorEntryID                 = 0x004C
	MAPIOriginalAuthorName                    = 0x004D
	MAPIOriginalSubmitTime                    = 0x004E
	MAPIReplyRecipientEntries                 = 0x004F
	MAPIReplyRecipientNames                   = 0x0050
	MAPIReceivedBySearchKey                   = 0x0051
	MAPIRcvdRepresentingSearchKey             = 0x0052
	MAPIReadReceiptSearchKey                  = 0x0053
	MAPIReportSearchKey                       = 0x0054
	MAPIOriginalDeliveryTime                  = 0x0055
	MAPIOriginalAuthorSearchKey               = 0x0056
	MAPIMessageToMe                           = 0x0057
	MAPIMessageCcMe                           = 0x0058
	MAPIMessageRecipMe                        = 0x0059
	MAPIOriginalSenderName                    = 0x005A
	MAPIOriginalSenderEntryID                 = 0x005B
	MAPIOriginalSenderSearchKey               = 0x005C
	MAPIOriginalSentRepresentingName          = 0x005D
	MAPIOriginalSentRepresentingEntryID       = 0x005E
	MAPIOriginalSentRepresentingSearchKey     = 0x005F
	MAPIStartDate                             = 0x0060
	MAPIEndDate                               = 0x0061
	MAPIOwnerApptID                           = 0x0062
	MAPIResponseRequested                     = 0x0063
	MAPISentRepresentingAddrtype              = 0x0064
	MAPISentRepresentingEmailAddress          = 0x0065
	MAPIOriginalSenderAddrtype                = 0x0066
	MAPIOriginalSenderEmailAddress            = 0x0067
	MAPIOriginalSentRepresentingAddrtype      = 0x0068
	MAPIOriginalSentRepresentingEmailAddress  = 0x0069
	MAPIConversationTopic                     = 0x0070
	MAPIConversationIndex                     = 0x0071
	MAPIOriginalDisplayBcc                    = 0x0072
	MAPIOriginalDisplayCc                     = 0x0073
	MAPIOriginalDisplayTo                     = 0x0074
	MAPIReceivedByAddrtype                    = 0x0075
	MAPIReceivedByEmailAddress                = 0x0076
	MAPIRcvdRepresentingAddrtype              = 0x0077
	MAPIRcvdRepresentingEmailAddress          = 0x0078
	MAPIOriginalAuthorAddrtype                = 0x0079
	MAPIOriginalAuthorEmailAddress            = 0x007A
	MAPIOriginallyIntendedRecipAddrtype       = 0x007B
	MAPIOriginallyIntendedRecipEmailAddress   = 0x007C
	MAPITransportMessageHeaders               = 0x007D
	MAPIDelegation                            = 0x007E
	MAPITnefCorrelationKey                    = 0x007F
	MAPIBody                                  = 0x1000
	MAPIBodyHTML                              = 0x1013
	MAPIReportText                            = 0x1001
	MAPIOriginatorAndDlExpansionHistory       = 0x1002
	MAPIReportingDlName                       = 0x1003
	MAPIReportingMtaCertificate               = 0x1004
	MAPIRtfSyncBodyCrc                        = 0x1006
	MAPIRtfSyncBodyCount                      = 0x1007
	MAPIRtfSyncBodyTag                        = 0x1008
	MAPIRtfCompressed                         = 0x1009
	MAPIRtfSyncPrefixCount                    = 0x1010
	MAPIRtfSyncTrailingCount                  = 0x1011
	MAPIOriginallyIntendedRecipEntryID        = 0x1012
	MAPIContentIntegrityCheck                 = 0x0C00
	MAPIExplicitConversion                    = 0x0C01
	MAPIIpmReturnRequested                    = 0x0C02
	MAPIMessageToken                          = 0x0C03
	MAPINdrReasonCode                         = 0x0C04
	MAPINdrDiagCode                           = 0x0C05
	MAPINonReceiptNotificationRequested       = 0x0C06
	MAPIDeliveryPoint                         = 0x0C07
	MAPIOriginatorNonDeliveryReportRequested  = 0x0C08
	MAPIOriginatorRequestedAlternateRecipient = 0x0C09
	MAPIPhysicalDeliveryBureauFaxDelivery     = 0x0C0A
	MAPIPhysicalDeliveryMode                  = 0x0C0B
	MAPIPhysicalDeliveryReportRequest         = 0x0C0C
	MAPIPhysicalForwardingAddress             = 0x0C0D
	MAPIPhysicalForwardingAddressRequested    = 0x0C0E
	MAPIPhysicalForwardingProhibited          = 0x0C0F
	MAPIPhysicalRenditionAttributes           = 0x0C10
	MAPIProofOfDelivery                       = 0x0C11
	MAPIProofOfDeliveryRequested              = 0x0C12
	MAPIRecipientCertificate                  = 0x0C13
	MAPIRecipientNumberForAdvice              = 0x0C14
	MAPIRecipientType                         = 0x0C15
	MAPIRegisteredMailType                    = 0x0C16
	MAPIReplyRequested                        = 0x0C17
	MAPIRequestedDeliveryMethod               = 0x0C18
	MAPISenderEntryID                         = 0x0C19
	MAPISenderName                            = 0x0C1A
	MAPISupplementaryInfo                     = 0x0C1B
	MAPITypeOfMtsUser                         = 0x0C1C
	MAPISenderSearchKey                       = 0x0C1D
	MAPISenderAddrtype                        = 0x0C1E
	MAPISenderEmailAddress                    = 0x0C1F
	MAPICurrentVersion                        = 0x0E00
	MAPIDeleteAfterSubmit                     = 0x0E01
	MAPIDisplayBcc                            = 0x0E02
	MAPIDisplayCc                             = 0x0E03
	MAPIDisplayTo                             = 0x0E04
	MAPIParentDisplay                         = 0x0E05
	MAPIMessageDeliveryTime                   = 0x0E06
	MAPIMessageFlags                          = 0x0E07
	MAPIMessageSize                           = 0x0E08
	MAPIParentEntryID                         = 0x0E09
	MAPISentmailEntryID                       = 0x0E0A
	MAPICorrelate                             = 0x0E0C
	MAPICorrelateMtsID                        = 0x0E0D
	MAPIDiscreteValues                        = 0x0E0E
	MAPIResponsibility                        = 0x0E0F
	MAPISpoolerStatus                         = 0x0E10
	MAPITransportStatus                       = 0x0E11
	MAPIMessageRecipients                     = 0x0E12
	MAPIMessageAttachments                    = 0x0E13
	MAPISubmitFlags                           = 0x0E14
	MAPIRecipientStatus                       = 0x0E15
	MAPITransportKey                          = 0x0E16
	MAPIMsgStatus                             = 0x0E17
	MAPIMessageDownloadTime                   = 0x0E18
	MAPICreationVersion                       = 0x0E19
	MAPIModifyVersion                         = 0x0E1A
	MAPIHasattach                             = 0x0E1B
	MAPIBodyCrc                               = 0x0E1C
	MAPINormalizedSubject                     = 0x0E1D
	MAPIRtfInSync                             = 0x0E1F
	MAPIAttachSize                            = 0x0E20
	MAPIAttachNum                             = 0x0E21
	MAPIPreprocess                            = 0x0E22
	MAPIOriginatingMtaCertificate             = 0x0E25
	MAPIProofOfSubmission                     = 0x0E26
	MAPIEntryID                               = 0x0FFF
	MAPIObjectType                            = 0x0FFE
	MAPIIcon                                  = 0x0FFD
	MAPIMiniIcon                              = 0x0FFC
	MAPIStoreEntryID                          = 0x0FFB
	MAPIStoreRecordKey                        = 0x0FFA
	MAPIRecordKey                             = 0x0FF9
	MAPIMappingSignature                      = 0x0FF8
	MAPIAccessLevel                           = 0x0FF7
	MAPIInstanceKey                           = 0x0FF6
	MAPIRowType                               = 0x0FF5
	MAPIAccess                                = 0x0FF4
	MAPIRowID                                 = 0x3000
	MAPIDisplayName                           = 0x3001
	MAPIAddrtype                              = 0x3002
	MAPIEmailAddress                          = 0x3003
	MAPIComment                               = 0x3004
	MAPIDepth                                 = 0x3005
	MAPIProviderDisplay                       = 0x3006
	MAPICreationTime                          = 0x3007
	MAPILastModificationTime                  = 0x3008
	MAPIResourceFlags                         = 0x3009
	MAPIProviderDllName                       = 0x300A
	MAPISearchKey                             = 0x300B
	MAPIProviderUID                           = 0x300C
	MAPIProviderOrdinal                       = 0x300D
	MAPIFormVersion                           = 0x3301
	MAPIFormClsid                             = 0x3302
	MAPIFormContactName                       = 0x3303
	MAPIFormCategory                          = 0x3304
	MAPIFormCategorySub                       = 0x3305
	MAPIFormHostMap                           = 0x3306
	MAPIFormHidden                            = 0x3307
	MAPIFormDesignerName                      = 0x3308
	MAPIFormDesignerGuID                      = 0x3309
	MAPIFormMessageBehavior                   = 0x330A
	MAPIDefaultStore                          = 0x3400
	MAPIStoreSupportMask                      = 0x340D
	MAPIStoreState                            = 0x340E
	MAPIIpmSubtreeSearchKey                   = 0x3410
	MAPIIpmOutboxSearchKey                    = 0x3411
	MAPIIpmWastebasketSearchKey               = 0x3412
	MAPIIpmSentmailSearchKey                  = 0x3413
	MAPIMdbProvider                           = 0x3414
	MAPIReceiveFolderSettings                 = 0x3415
	MAPIValidFolderMask                       = 0x35DF
	MAPIIpmSubtreeEntryID                     = 0x35E0
	MAPIIpmOutboxEntryID                      = 0x35E2
	MAPIIpmWastebasketEntryID                 = 0x35E3
	MAPIIpmSentmailEntryID                    = 0x35E4
	MAPIViewsEntryID                          = 0x35E5
	MAPICommonViewsEntryID                    = 0x35E6
	MAPIFinderEntryID                         = 0x35E7
	MAPIContainerFlags                        = 0x3600
	MAPIFolderType                            = 0x3601
	MAPIContentCount                          = 0x3602
	MAPIContentUnread                         = 0x3603
	MAPICreateTemplates                       = 0x3604
	MAPIDetailsTable                          = 0x3605
	MAPISearch                                = 0x3607
	MAPISelectable                            = 0x3609
	MAPISubfolders                            = 0x360A
	MAPIStatus                                = 0x360B
	MAPIAnr                                   = 0x360C
	MAPIContentsSortOrder                     = 0x360D
	MAPIContainerHierarchy                    = 0x360E
	MAPIContainerContents                     = 0x360F
	MAPIFolderAssociatedContents              = 0x3610
	MAPIDefCreateDl                           = 0x3611
	MAPIDefCreateMailuser                     = 0x3612
	MAPIContainerClass                        = 0x3613
	MAPIContainerModifyVersion                = 0x3614
	MAPIAbProviderID                          = 0x3615
	MAPIDefaultViewEntryID                    = 0x3616
	MAPIAssocContentCount                     = 0x3617
	MAPIAttachmentX400Parameters              = 0x3700
	MAPIAttachDataObj                         = 0x3701
	MAPIAttachEncoding                        = 0x3702
	MAPIAttachExtension                       = 0x3703
	MAPIAttachFilename                        = 0x3704
	MAPIAttachMethod                          = 0x3705
	MAPIAttachLongFilename                    = 0x3707
	MAPIAttachPathname                        = 0x3708
	MAPIAttachRendering                       = 0x3709
	MAPIAttachTag                             = 0x370A
	MAPIRenderingPosition                     = 0x370B
	MAPIAttachTransportName                   = 0x370C
	MAPIAttachLongPathname                    = 0x370D
	MAPIAttachMimeTag                         = 0x370E
	MAPIAttachAdditionalInfo                  = 0x370F
	MAPIDisplayType                           = 0x3900
	MAPITemplateID                            = 0x3902
	MAPIPrimaryCapability                     = 0x3904
	MAPI7bitDisplayName                       = 0x39FF
	MAPIAccount                               = 0x3A00
	MAPIAlternateRecipient                    = 0x3A01
	MAPICallbackTelephoneNumber               = 0x3A02
	MAPIConversionProhibited                  = 0x3A03
	MAPIDiscloseRecipients                    = 0x3A04
	MAPIGeneration                            = 0x3A05
	MAPIGivenName                             = 0x3A06
	MAPIGovernmentIDNumber                    = 0x3A07
	MAPIBusinessTelephoneNumber               = 0x3A08
	MAPIHomeTelephoneNumber                   = 0x3A09
	MAPIInitials                              = 0x3A0A
	MAPIKeyword                               = 0x3A0B
	MAPILanguage                              = 0x3A0C
	MAPILocation                              = 0x3A0D
	MAPIMailPermission                        = 0x3A0E
	MAPIMhsCommonName                         = 0x3A0F
	MAPIOrganizationalIDNumber                = 0x3A10
	MAPISurname                               = 0x3A11
	MAPIOriginalEntryID                       = 0x3A12
	MAPIOriginalDisplayName                   = 0x3A13
	MAPIOriginalSearchKey                     = 0x3A14
	MAPIPostalAddress                         = 0x3A15
	MAPICompanyName                           = 0x3A16
	MAPITitle                                 = 0x3A17
	MAPIDepartmentName                        = 0x3A18
	MAPIOfficeLocation                        = 0x3A19
	MAPIPrimaryTelephoneNumber                = 0x3A1A
	MAPIBusiness2TelephoneNumber              = 0x3A1B
	MAPIMobileTelephoneNumber                 = 0x3A1C
	MAPIRadioTelephoneNumber                  = 0x3A1D
	MAPICarTelephoneNumber                    = 0x3A1E
	MAPIOtherTelephoneNumber                  = 0x3A1F
	MAPITransmitableDisplayName               = 0x3A20
	MAPIPagerTelephoneNumber                  = 0x3A21
	MAPIUserCertificate                       = 0x3A22
	MAPIPrimaryFaxNumber                      = 0x3A23
	MAPIBusinessFaxNumber                     = 0x3A24
	MAPIHomeFaxNumber                         = 0x3A25
	MAPICountry                               = 0x3A26
	MAPILocality                              = 0x3A27
	MAPIStateOrProvince                       = 0x3A28
	MAPIStreetAddress                         = 0x3A29
	MAPIPostalCode                            = 0x3A2A
	MAPIPostOfficeBox                         = 0x3A2B
	MAPITelexNumber                           = 0x3A2C
	MAPIIsdnNumber                            = 0x3A2D
	MAPIAssistantTelephoneNumber              = 0x3A2E
	MAPIHome2TelephoneNumber                  = 0x3A2F
	MAPIAssistant                             = 0x3A30
	MAPISendRichInfo                          = 0x3A40
	MAPIWeddingAnniversary                    = 0x3A41
	MAPIBirthday                              = 0x3A42
	MAPIHobbies                               = 0x3A43
	MAPIMiddleName                            = 0x3A44
	MAPIDisplayNamePrefix                     = 0x3A45
	MAPIProfession                            = 0x3A46
	MAPIPreferredByName                       = 0x3A47
	MAPISpouseName                            = 0x3A48
	MAPIComputerNetworkName                   = 0x3A49
	MAPICustomerID                            = 0x3A4A
	MAPITtytddPhoneNumber                     = 0x3A4B
	MAPIFtpSite                               = 0x3A4C
	MAPIGender                                = 0x3A4D
	MAPIManagerName                           = 0x3A4E
	MAPINickname                              = 0x3A4F
	MAPIPersonalHomePage                      = 0x3A50
	MAPIBusinessHomePage                      = 0x3A51
	MAPIContactVersion                        = 0x3A52
	MAPIContactEntryids                       = 0x3A53
	MAPIContactAddrtypes                      = 0x3A54
	MAPIContactDefaultAddressIndex            = 0x3A55
	MAPIContactEmailAddresses                 = 0x3A56
	MAPICompanyMainPhoneNumber                = 0x3A57
	MAPIChildrensNames                        = 0x3A58
	MAPIHomeAddressCity                       = 0x3A59
	MAPIHomeAddressCountry                    = 0x3A5A
	MAPIHomeAddressPostalCode                 = 0x3A5B
	MAPIHomeAddressStateOrProvince            = 0x3A5C
	MAPIHomeAddressStreet                     = 0x3A5D
	MAPIHomeAddressPostOfficeBox              = 0x3A5E
	MAPIOtherAddressCity                      = 0x3A5F
	MAPIOtherAddressCountry                   = 0x3A60
	MAPIOtherAddressPostalCode                = 0x3A61
	MAPIOtherAddressStateOrProvince           = 0x3A62
	MAPIOtherAddressStreet                    = 0x3A63
	MAPIOtherAddressPostOfficeBox             = 0x3A64
	MAPIStoreProviders                        = 0x3D00
	MAPIAbProviders                           = 0x3D01
	MAPITransportProviders                    = 0x3D02
	MAPIDefaultProfile                        = 0x3D04
	MAPIAbSearchPath                          = 0x3D05
	MAPIAbDefaultDir                          = 0x3D06
	MAPIAbDefaultPab                          = 0x3D07
	MAPIFilteringHooks                        = 0x3D08
	MAPIServiceName                           = 0x3D09
	MAPIServiceDllName                        = 0x3D0A
	MAPIServiceEntryName                      = 0x3D0B
	MAPIServiceUID                            = 0x3D0C
	MAPIServiceExtraUids                      = 0x3D0D
	MAPIServices                              = 0x3D0E
	MAPIServiceSupportFiles                   = 0x3D0F
	MAPIServiceDeleteFiles                    = 0x3D10
	MAPIAbSearchPathUpdate                    = 0x3D11
	MAPIProfileName                           = 0x3D12
	MAPIIdentityDisplay                       = 0x3E00
	MAPIIdentityEntryID                       = 0x3E01
	MAPIResourceMethods                       = 0x3E02
	MAPIResourceType                          = 0x3E03
	MAPIStatusCode                            = 0x3E04
	MAPIIdentitySearchKey                     = 0x3E05
	MAPIOwnStoreEntryID                       = 0x3E06
	MAPIResourcePath                          = 0x3E07
	MAPIStatusString                          = 0x3E08
	MAPIX400DeferredDeliveryCancel            = 0x3E09
	MAPIHeaderFolderEntryID                   = 0x3E0A
	MAPIRemoteProgress                        = 0x3E0B
	MAPIRemoteProgressText                    = 0x3E0C
	MAPIRemoteValidateOk                      = 0x3E0D
	MAPIControlFlags                          = 0x3F00
	MAPIControlStructure                      = 0x3F01
	MAPIControlType                           = 0x3F02
	MAPIDeltax                                = 0x3F03
	MAPIDeltay                                = 0x3F04
	MAPIXpos                                  = 0x3F05
	MAPIYpos                                  = 0x3F06
	MAPIControlID                             = 0x3F07
	MAPIInitialDetailsPane                    = 0x3F08
	MAPIIdSecureMin                           = 0x67F0
	MAPIIdSecureMax                           = 0x67FF
)
