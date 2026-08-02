// vybe-test: kotlin/kotlin_sequences_generate/test_iterator_style_consumption_chain
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequences_generate.rs

fun main() {
            val source = sequenceOf("a", "bb", "c").iterator()
            var out = ""
            while (source.hasNext()) {
                out += source.next()
            }
            println(out)
        }

