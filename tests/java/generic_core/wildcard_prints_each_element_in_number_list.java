// vybe-test: java/generic_core/wildcard_prints_each_element_in_number_list
// origin: languages/java/tests/java/test_generic_core.rs

public class Main {
static class Dump {
            static void dump(java.util.List<? extends Number> nums) {
                for (Number n : nums) System.out.println(n.intValue());
            }
        }
    public static void main(String[] args) {
java.util.List<Integer> nums = java.util.Arrays.asList(2, 4); Dump.dump(nums);
    }
}

