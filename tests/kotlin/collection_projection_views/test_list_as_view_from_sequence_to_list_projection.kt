// vybe-test: kotlin/collection_projection_views/test_list_as_view_from_sequence_to_list_projection
// origin: languages/kotlin/tests/kotlin/test_collection_projection_views.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = sequenceOf(1, 2, 3, 4)
            val view = seq.toList()
            val rev = view.asReversed()
            __check((rev.joinToString(",")).toString(), "4,3,2,1")
            __check((view[0]).toString(), "1")
            __check((view.last()).toString(), "4")
        }
