void greet(String name, {String prefix = "Hello", int times = 1}) {
    for (var i = 0; i < times; i++) {
        print("$prefix, $name!");
    }
}

int? maybeNull(int? value) {
    return DateTime.now().microsecondsSinceEpoch < 0 ? null : value;
}

void main() {
    print("--- Null-Aware Tests ---");
    int? x;
    print("x ?? 5: ${x ?? 5}"); // Should be 5
    
    int? y = DateTime.now().millisecondsSinceEpoch > 0 ? 10 : null;
    print("y ?? 5: ${y ?? 5}"); // Should be 10
    
    // Null-aware assignment
    int? z;
    z ??= 20;
    print("z ??= 20: $z"); // Should be 20
    z = maybeNull(z);
    z ??= 30;
    print("z ??= 30 (already 20): $z"); // Should be 20
    
    // Null-safe access
    String? s = DateTime.now().millisecondsSinceEpoch > 0 ? "hello" : null;
    print("s?.length: ${s?.length}"); // Should be 5
    
    String? s2 = DateTime.now().millisecondsSinceEpoch % 2 == 0 ? null : "world";
    print("s2?.length: ${s2?.length}"); // Should be null
    
    print("\n--- Named Parameter Tests ---");
    greet("World"); // Default: Hello, World!
    greet("Vybe", prefix: "Welcome"); // Welcome, Vybe!
    greet("Repeater", times: 2); // Hello, Repeater! x2
    
    print("\n--- List/Map/Set Literals ---");
    var list = [1, 2, 3];
    print("List: $list");
    
    var map = {"a": 1, "b": 2};
    print("Map: $map");
    
    var set = {1, 2, 3};
    print("Set: $set");
}
