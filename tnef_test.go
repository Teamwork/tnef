package tnef

import (
	"errors"
	"testing"
)

func TestAttachments(t *testing.T) {
	tnef, err := DecodeFile("./testfiles/attachments.dat")
	if err != nil {
		t.Error(err)
		return
	}

	if len(tnef.Attachments) != 2 {
		t.Error(errors.New("The decoded file should have two attachments"))
		return
	}

	for _, a := range tnef.Attachments {
		if len(a.Data) == 0 {
			t.Error(errors.New("Attachment " + a.Title + " has no data in it!"))
		}
	}
}

func TestBadChecksum(t *testing.T) {
	_, err := DecodeFile("./testfiles/badchecksum.dat")
	if err == nil {
		t.Error(errors.New("This should fail because of a bad checksum."))
	}
}
