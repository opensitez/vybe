// vybe-test: kotlin/strings/test_string_index_and_iteration
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun main() {
            val word = "chat"
            var letters = ""
            for (ch in word) {
                letters += ch
            }
            println(word[0])
            println(word[3])
            println(letters)
        }

