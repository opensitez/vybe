// vybe-test: kotlin/annotations/test_annotation_with_enum_argument
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

enum class Level { LOW, MEDIUM, HIGH }

        annotation class Tier(val level: Level)

        @Tier(Level.HIGH)
        fun service() = "tiered"

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((service()).toString(), "tiered")
        }
