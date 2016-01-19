package tnef

import (
	"errors"
	"io/ioutil"
	"strings"
)

const (
	TNEF_SIGNATURE = 0x223e9f78
	LVL_MESSAGE    = 0x01
	LVL_ATTACHMENT = 0x02

	ATTOWNER                   = 0x0000 // Owner
	ATTSENTFOR                 = 0x0001 // Sent For
	ATTDELEGATE                = 0x0002 // Delegate
	ATTDATESTART               = 0x0006 // Date Start
	ATTDATEEND                 = 0x0007 // Date End
	ATTAIDOWNER                = 0x0008 // Owner Appointment ID
	ATTREQUESTRES              = 0x0009 // Response Requested.
	ATTFROM                    = 0x8000 // From
	ATTSUBJECT                 = 0x8004 // Subject
	ATTDATESENT                = 0x8005 // Date Sent
	ATTDATERECD                = 0x8006 // Date Recieved
	ATTMESSAGESTATUS           = 0x8007 // Message Status
	ATTMESSAGECLASS            = 0x8008 // Message Class
	ATTMESSAGEID               = 0x8009 // Message ID
	ATTPARENTID                = 0x800a // Parent ID
	ATTCONVERSATIONID          = 0x800b // Conversation ID
	ATTBODY                    = 0x800c // Body
	ATTPRIORITY                = 0x800d // Priority
	ATTATTACHDATA              = 0x800f // Attachment Data
	ATTATTACHTITLE             = 0x8010 // Attachment File Name
	ATTATTACHMETAFILE          = 0x8011 // Attachment Meta File
	ATTATTACHCREATEDATE        = 0x8012 // Attachment Creation Date
	ATTATTACHMODIFYDATE        = 0x8013 // Attachment Modification Date
	ATTDATEMODIFY              = 0x8020 // Date Modified
	ATTATTACHTRANSPORTFILENAME = 0x9001 // Attachment Transport Filename
	ATTATTACHRENDDATA          = 0x9002 // Attachment Rendering Data
	ATTMAPIPROPS               = 0x9003 // MAPI Properties
	ATTRECIPTABLE              = 0x9004 // Recipients
	ATTATTACHMENT              = 0x9005 // Attachment
	ATTTNEFVERSION             = 0x9006 // TNEF Version
	ATTOEMCODEPAGE             = 0x9007 // OEM Codepage
	ATTORIGNINALMESSAGECLASS   = 0x9008 // Original Message Class
)

type TNEFObject struct {
	Level  int
	Name   int
	Type   int
	Data   []byte
	Length int
}

type TNEFAttachment struct {
	Title string
	Data  []byte
}

func (a *TNEFAttachment) AddAttr(obj TNEFObject) {
	if obj.Name == ATTATTACHTITLE {
		a.Title = strings.Replace(string(obj.Data), "\x00", "", -1)
	} else if obj.Name == ATTATTACHDATA {
		a.Data = obj.Data
	}
}

func DecodeFile(path string) ([]*TNEFAttachment, error) {
	data, err := ioutil.ReadFile(path)
	if err != nil {
		return nil, err
	}

	return Decode(data)
}

func Decode(data []byte) ([]*TNEFAttachment, error) {
	if byte_to_int(data[0:4]) != TNEF_SIGNATURE {
		return nil, errors.New("Signature didn't match valid TNEF file")
	}

	//key := binary.LittleEndian.Uint32(data[4:6])
	offset := 6
	var attachments []*TNEFAttachment
	var attachment *TNEFAttachment

	for offset < len(data) {
		obj := decodeTNEFObject(data[offset:])
		offset += obj.Length

		if obj.Name == ATTATTACHRENDDATA {
			attachment = new(TNEFAttachment)
			attachments = append(attachments, attachment)
		} else if obj.Level == LVL_ATTACHMENT {
			attachment.AddAttr(obj)
		} else if obj.Name == ATTMAPIPROPS {
			// TODO
		}
	}

	return attachments, nil
}

func decodeTNEFObject(data []byte) (object TNEFObject) {
	offset := 0

	object.Level = byte_to_int(data[offset : offset+1])
	offset += 1
	object.Name = byte_to_int(data[offset : offset+2])
	offset += 2
	object.Type = byte_to_int(data[offset : offset+2])
	offset += 2
	att_length := byte_to_int(data[offset : offset+4])
	offset += 4
	object.Data = data[offset : offset+att_length]
	offset += att_length
	//checksum := byte_to_int(data[offset : offset+2])
	offset += 2

	object.Length = offset
	return
}
