![Godoc](https://camo.githubusercontent.com/a7c641a533908ef24c4e42195fa72c6fcd2ae1f0/68747470733a2f2f676f646f632e6f72672f6769746875622e636f6d2f507565726b69746f42696f2f676f71756572793f7374617475732e706e67)
## Go library to extract body and attachments from TNEF files
With this library you can extract the body and attachments from Transport Neutral Encapsulation Format (TNEF) files. This work is based on https://github.com/koodaamo/tnefparse.

## Example usage
```go
package main

import (
	"io/ioutil"
	"os"

	"github.com/teamwork/tnef"
)

func main() {
	t, err := tnef.DecodeFile("./winmail.dat")
	if err != nil {
		return
	}

	wd, _ := os.Getwd()
	for _, a := range t.Attachments {
		ioutil.WriteFile(wd+"/"+a.Title, a.Data, 0777)
	}

	ioutil.WriteFile(wd+"/bodyHTML.html", t.BodyHTML, 0777)
	ioutil.WriteFile(wd+"/bodyPlain.html", t.Body, 0777)
}

```
