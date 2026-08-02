// vybe-test: kotlin/companion_objects/test_companion_object_can_implement_comparator_for_custom_sorting
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

data class Entry(val value: Int)

        class Holder {
            companion object : Comparator<Entry> {
                override fun compare(left: Entry, right: Entry): Int {
                    return right.value - left.value
                }
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(Entry(1), Entry(3), Entry(2))
            val sorted = values.sortedWith(Holder.Companion)
            __check((sorted.joinToString(",") { it.value.toString() }).toString(), "3,2,1")
        }
