// vybe-test: kotlin/smart_casts/test_looping_type_checks_in_collection
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

fun main() {
            val values = listOf<Any>(1, "two", 3, "four")
            var strings = 0
            var totalLen = 0
            for (item in values) {
                if (item is String) {
                    strings += 1
                    totalLen += item.length
                }
            }
            println(strings)
            println(totalLen)
        }

