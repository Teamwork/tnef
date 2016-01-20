package tnef

func byte_to_int(data []byte) int {
	var num int
	var n uint
	for _, b := range data {
		num += (int(b) << n)
		n += 8
	}
	return num
}

func min(a, b int) int {
	if a > b {
		return b
	}
	return a
}
