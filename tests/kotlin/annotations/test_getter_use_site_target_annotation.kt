// vybe-test: kotlin/annotations/test_getter_use_site_target_annotation
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

class Record {
            @get:Deprecated("legacy")
            val label: String = "label"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = Record()
            __check((r.label).toString(), "label")
        }
