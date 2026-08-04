// vybe-test: java/anonymous_classes/anonymous_class_nested_in_method
// origin: languages/java/tests/java/test_anonymous_classes.rs

public class Main {
static class Maker {
            Runnable make() {
                return new Runnable() { public void run() { System.out.println("nested"); } };
            }
        }
    public static void main(String[] args) {
Maker m = new Maker(); m.make().run();
    }
}

