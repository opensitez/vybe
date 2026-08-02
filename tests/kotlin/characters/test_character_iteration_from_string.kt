// vybe-test: kotlin/characters/test_character_iteration_from_string
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun main() {
            var out = ""
            for (c in "kotlin") {
                out += c
            }
            val first = "kotlin"[0]
            val last = "kotlin"["kotlin".lastIndex]
            println(out)
            println(first)
            println(last)
        }

