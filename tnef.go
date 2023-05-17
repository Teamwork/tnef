// Package tnef extracts the body and attachments from Microsoft TNEF files.
package tnef // import "github.com/teamwork/tnef"

import (
	"errors"
	"io/ioutil"
	"strings"
)

const (
	tnefSignature = 0x223e9f78
	//lvlMessage    = 0x01
	lvlAttachment = 0x02
)

// These can be used to figure out the type of attribute
// an object is
const (
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
	ATTDATERECD                = 0x8006 // Date Received
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

type tnefObject struct {
	Level  int
	Name   int
	Type   int
	Data   []byte
	Length int
}

// Attachment contains standard attachments that are embedded
// within the TNEF file, with the name and data of the file extracted.
type Attachment struct {
	Title string
	Data  []byte
}

// ErrNoMarker signals that the file did not start with the fixed TNEF marker,
// meaning it's not in the TNEF file format we recognize (e.g. it just has the
// .tnef extension, or a wrong MIME type).
var ErrNoMarker = errors.New("file did not begin with a TNEF marker")

// Data contains the various data from the extracted TNEF file.
type Data struct {
	Body        []byte
	BodyHTML    []byte
	Attachments []*Attachment
	Attributes  []MAPIAttribute
}

func (a *Attachment) addAttr(obj tnefObject) {
	switch obj.Name {
	case ATTATTACHTITLE:
		a.Title = strings.Replace(string(obj.Data), "\x00", "", -1)
	case ATTATTACHDATA:
		a.Data = obj.Data
	}
}

// DecodeFile is a utility function that reads the file into memory
// before calling the normal Decode function on the data.
func DecodeFile(path string) (*Data, error) {
	data, err := ioutil.ReadFile(path)
	if err != nil {
		return nil, err
	}

	return Decode(data)
}

// Decode will accept a stream of bytes in the TNEF format and extract the
// attachments and body into a Data object.
func Decode(data []byte) (*Data, error) {
	if len(data) < 4 || byteToInt(data[0:4]) != tnefSignature {
		return nil, ErrNoMarker
	}

	//key := binary.LittleEndian.Uint32(data[4:6])
	offset := 6
	var attachment *Attachment
	tnef := &Data{
		Attachments: []*Attachment{},
	}

	for offset < len(data) {
		obj := decodeTNEFObject(data[offset:])
		offset += obj.Length

		if obj.Name == ATTATTACHRENDDATA {
			attachment = new(Attachment)
			tnef.Attachments = append(tnef.Attachments, attachment)
		} else if obj.Level == lvlAttachment {
			attachment.addAttr(obj)
		} else if obj.Name == ATTMAPIPROPS {
			var err error
			tnef.Attributes, err = decodeMapi(obj.Data)
			if err != nil {
				return nil, err
			}

			// Get the body property if it's there
			for _, attr := range tnef.Attributes {
				switch attr.Name {
				case MAPIBody:
					tnef.Body = attr.Data
				case MAPIBodyHTML:
					tnef.BodyHTML = attr.Data
				}
			}
		}
	}

	return tnef, nil
}

func decodeTNEFObject(data []byte) (object tnefObject) {
	offset := 0

	object.Level = byteToInt(data[offset : offset+1])
	offset++
	object.Name = byteToInt(data[offset : offset+2])
	offset += 2
	object.Type = byteToInt(data[offset : offset+2])
	offset += 2
	attLength := byteToInt(data[offset : offset+4])
	offset += 4
	object.Data = data[offset : offset+attLength]
	offset += attLength
	//checksum := byteToInt(data[offset : offset+2])
	offset += 2

	object.Length = offset
	return
}
