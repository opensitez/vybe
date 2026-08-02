// vybe-test: kotlin/strings/test_to_list_of_chars_roundtrip
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun main() {
            val chars = "dog".toCharArray()
            println(chars.joinToString(","))
            var rebuilt = ""
            for (ch in chars) {
                rebuilt += ch
            }
            println(rebuilt)
        }

