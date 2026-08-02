// vybe-test: kotlin/destructuring/test_destructuring_from_map_entries
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun main() {
            val inventory = mapOf("apple" to 2, "orange" to 5)
            var labels = ""
            var quantities = 0
            for ((name, count) in inventory) {
                labels += name[0].toString()
                quantities += count
            }
            println(labels)
            println(quantities)
        }

