// vybe-test: kotlin/interfaces/test_interface_array_dispatch
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Score { fun score(): Int }
        class A : Score { override fun score(): Int = 1 }
        class B : Score { override fun score(): Int = 2 }
        class C : Score { override fun score(): Int = 3 }

        fun total(items: Array<Score>): Int {
            var sum = 0
            for (item in items) {
                sum += item.score()
            }
            return sum
        }

        fun main() {
            println(total(arrayOf(A(), B(), C())))
        }

