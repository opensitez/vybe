// vybe-test: kotlin/enums/test_enum_class_matching
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Status {
            PENDING, APPROVED, REJECTED
        }

        fun main() {
            val s = Status.APPROVED
            if (s == 1) {
                println("Approved Status")
            } else {
                println("Other Status")
            }
        }

