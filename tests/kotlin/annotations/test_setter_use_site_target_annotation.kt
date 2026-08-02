// vybe-test: kotlin/annotations/test_setter_use_site_target_annotation
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

class Holder {
            @set:Suppress("UNUSED")
            var value: Int = 0
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val h = Holder()
            h.value = 6
            __check((h.value).toString(), "6")
        }
