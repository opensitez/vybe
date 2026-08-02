// vybe-test: kotlin/kotlin_progressions/test_char_progression
// origin: languages/kotlin/tests/kotlin/test_kotlin_progressions.rs

fun main() {
            var out = ""
            for (c in 'a'..'d') { out = out + c }
            println(out)
            println(('c' in 'a'..'d').toString())
            println(('x' in 'a'..'d').toString())
        }

