// vybe-test: kotlin/type_casts/test_is_filtering_in_loop
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun main() {
            val items: Array<Any?> = arrayOf(1, "x", true, 2.5, null)
            var count = 0
            var stringLen = 0
            var boolSeen = false
            for (item in items) {
                if (item is Int) {
                    count += item
                } else if (item is String) {
                    stringLen = item.length
                } else if (item is Boolean) {
                    boolSeen = item
                }
            }
            println(count)
            println(stringLen)
            println(boolSeen)
        }

