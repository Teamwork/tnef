package tnef

import (
	"testing"

	"github.com/teamwork/test"
	"github.com/teamwork/utils/sliceutil"
)

func TestAttachments(t *testing.T) {
	tests := []struct {
		in              string
		wantAttachments []string
		wantErr         string
	}{
		{"attachments", []string{
			"ZAPPA_~2.JPG",
			"bookmark.htm",
		}, ""},
		// will panic!
		//{"panic", []string{
		//	"ZAPPA_~2.JPG",
		//	"bookmark.htm",
		//}},
		//{"MAPI_ATTACH_DATA_OBJ", []string{
		//	"VIA_Nytt_1402.doc",
		//	"VIA_Nytt_1402.pdf",
		//	"VIA_Nytt_14021.htm",
		//	"MAPI_ATTACH_DATA_OBJ-body.rtf",
		//}},
		//{"MAPI_OBJECT", []string{
		//	"Untitled_Attachment",
		//	"MAPI_OBJECT-body.rtf",
		//}},
		//{"body", []string{
		//	"body-body.html",
		//}},
		//{"data-before-name", []string{
		//	"AUTOEXEC.BAT",
		//	"CONFIG.SYS",
		//	"boot.ini",
		//	"data-before-name-body.rtf",
		//}},
		// {"garbage-at-end", []string{}, ""}, // panics
		//{"long-filename", []string{
		//	"long-filename-body.rtf",
		//}},
		//{"missing-filenames", []string{
		//	"missing-filenames-body.rtf",
		//}},
		{"multi-name-property", []string{}, ""},
		//{"multi-value-attribute", []string{
		//	"208225__5_seconds__Voice_Mail.mp3",
		//	"multi-value-attribute-body.rtf",
		//}},
		{"one-file", []string{
			"AUTHORS",
		}, ""},
		//{"rtf", []string{
		//	"rtf-body.rtf",
		//}},
		//{"triples", []string{
		//	"triples-body.rtf",
		//}},
		{"two-files", []string{
			"AUTHORS",
			"README",
		}, ""},
		{"unicode-mapi-attr-name", []string{
			"spaconsole2.cfg",
			"image001.png",
			"image002.png",
			"image003.png",
		}, ""},
		{"unicode-mapi-attr", []string{
			"example.dat",
		}, ""},

		// Invalid files.
		{"badchecksum", nil, ErrNoMarker.Error()},
		{"empty-file", nil, ErrNoMarker.Error()},
	}

	for _, tt := range tests {
		t.Run(tt.in, func(t *testing.T) {
			out, err := Decode(test.Read(t, "./testdata", tt.in+".tnef"))
			if !test.ErrorContains(err, tt.wantErr) {
				t.Fatalf("wrong err\ngot:  %v\nwant: %v", err, tt.wantErr)
			}
			if err != nil {
				return
			}

			if len(out.Attachments) != len(tt.wantAttachments) {
				t.Errorf("wrong length; want %v, got %v",
					len(tt.wantAttachments), len(out.Attachments))
			}

			titles := []string{}
			for _, a := range out.Attachments {
				titles = append(titles, a.Title)
				//if len(a.Data) == 0 {
				//	t.Error("len(a.Data) is 0")
				//}
			}
			for _, want := range tt.wantAttachments {
				if !sliceutil.InStringSlice(titles, want) {
					t.Errorf("did not find %#v in the attachments: %#v", want, titles)
				}
			}
		})
	}
}
