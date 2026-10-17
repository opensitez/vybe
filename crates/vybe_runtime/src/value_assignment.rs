//! Rejected owned-assignment experiment, retained only for tests and measurement.

use crate::value::Value;

impl Value {
    #[inline(always)]
    pub(crate) fn assign_owned(&mut self, value: Self) {
        match (self, value) {
            (Self::I32(dst), Self::I32(value)) => *dst = value,
            (Self::I64(dst), Self::I64(value)) => *dst = value,
            (Self::F32(dst), Self::F32(value)) => *dst = value,
            (Self::F64(dst), Self::F64(value)) => *dst = value,
            (Self::V128(dst), Self::V128(value)) => *dst = value,
            (dst, value) => *dst = value,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::Arc;

    fn assert_same_numeric(actual: &Value, expected: &Value) {
        match (actual, expected) {
            (Value::I32(a), Value::I32(b)) => assert_eq!(a, b),
            (Value::I64(a), Value::I64(b)) => assert_eq!(a, b),
            (Value::F32(a), Value::F32(b)) => assert_eq!(a.to_bits(), b.to_bits()),
            (Value::F64(a), Value::F64(b)) => assert_eq!(a.to_bits(), b.to_bits()),
            (Value::V128(a), Value::V128(b)) => assert_eq!(a, b),
            (Value::Null, Value::Null) => {}
            _ => panic!("assignment changed the value variant"),
        }
    }

    #[test]
    fn owned_assignment_preserves_numeric_variants_and_exact_bits() {
        let values = [
            Value::Null,
            Value::I32(i32::MIN),
            Value::I32(i32::MAX),
            Value::I64(i64::MIN),
            Value::I64(i64::MAX),
            Value::F32(f32::from_bits(0x8000_0000)),
            Value::F32(f32::from_bits(0x7fc0_1234)),
            Value::F64(f64::from_bits(0x8000_0000_0000_0000)),
            Value::F64(f64::from_bits(0x7ff8_1234_5678_9abc)),
            Value::V128([0xa5; 16]),
        ];
        for initial in &values {
            for value in &values {
                let mut slot = initial.clone();
                slot.assign_owned(value.clone());
                assert_same_numeric(&slot, value);
            }
        }
    }

    #[test]
    fn owned_reference_fallback_moves_identity_and_releases_old_value() {
        let mut slot = Value::String(String::from("old").into());
        let old = match &slot {
            Value::String(text) => Arc::downgrade(text),
            _ => unreachable!(),
        };
        let incoming = Value::String(String::from("incoming").into());
        let alias = match &incoming {
            Value::String(text) => Arc::clone(text),
            _ => unreachable!(),
        };
        slot.assign_owned(incoming);
        assert!(old.upgrade().is_none());
        match &slot {
            Value::String(text) => assert!(Arc::ptr_eq(text, &alias)),
            _ => panic!("reference assignment changed the variant"),
        }
        assert_eq!(Arc::strong_count(&alias), 2);
        slot.assign_owned(Value::I32(7));
        assert_eq!(Arc::strong_count(&alias), 1);
        assert_same_numeric(&slot, &Value::I32(7));
    }

    #[test]
    #[ignore = "paired owned numeric assignment measurement, not VM timing"]
    fn paired_owned_assignment_benchmark() {
        use std::hint::black_box;
        use std::time::Instant;
        type Assigner = fn(&mut Value, Value);
        #[inline(never)]
        fn previous(dst: &mut Value, value: Value) { *dst = value; }
        #[inline(never)]
        fn candidate(dst: &mut Value, value: Value) { dst.assign_owned(value); }
        fn sample(assign: Assigner, slots: &mut [Value], passes: usize) -> f64 {
            let start = Instant::now();
            for pass in 0..passes {
                for (index, slot) in slots.iter_mut().enumerate() {
                    black_box(assign)(black_box(slot), Value::I32((index ^ pass) as i32));
                }
            }
            let elapsed = start.elapsed().as_secs_f64();
            for (index, slot) in slots.iter().enumerate() {
                assert!(matches!(slot, Value::I32(value) if *value == (index ^ (passes - 1)) as i32));
            }
            elapsed
        }
        for len in [4096usize, 65_536, 262_144] {
            let mut slots = vec![Value::I32(0); len];
            let passes = 1_048_576 / len;
            sample(previous, &mut slots, passes);
            sample(candidate, &mut slots, passes);
            let mut ratios = Vec::with_capacity(21);
            for pair in 0..21 {
                let mut times = [0.0; 2];
                for is_new in if pair % 2 == 0 { [false, true] } else { [true, false] } {
                    times[usize::from(is_new)] = sample(if is_new { candidate } else { previous }, &mut slots, passes);
                }
                assert!(times.iter().all(|time| time.is_finite() && *time > 0.0));
                ratios.push(times[1] / times[0]);
            }
            ratios.sort_by(f64::total_cmp);
            println!("owned I32 assignment slots={len} candidate/previous={:.6}; assignment helper only, allocation and validation excluded", ratios[10]);
        }
    }
}

