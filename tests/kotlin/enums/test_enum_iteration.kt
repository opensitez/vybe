// vybe-test: kotlin/enums/test_enum_iteration
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Light { RED, GREEN, BLUE }

        fun main() {
            var count = 0
            for (entry in arrayOf(Light.RED, Light.GREEN, Light.BLUE)) {
                if (entry == Light.RED) {
                    count += 1
                } else if (entry == Light.GREEN) {
                    count += 1
                } else {
                    count += 1
                }
            }
            println(count)
        }

