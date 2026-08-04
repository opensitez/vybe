// vybe-test: dart/enums_advanced/enhanced_enum_with_field
// origin: languages/dart/tests/dart/test_enums_advanced.rs

enum Planet {
  mercury(3.303e+23, 2.4397e6),
  venus(4.869e+24, 6.0518e6);

  final double mass;
  final double radius;
  const Planet(this.mass, this.radius);
}

void main() {}
