//! Arbitrary-precision integer — the backing for `Value::BigInt`.
//!
//! ECMA-262 §6.1.6.2: BigInt operations are defined on **mathematical
//! integers** — there is no overflow and no wrap at any width. This is
//! the host value type js-primitive-builtins models as opaque `bigint`;
//! the only place a 64-bit wrap is ever legal is the wasm boundary,
//! where the js-types JS-API prescribes ToBigInt64 / ToBigUint64
//! (`to_i64_wrapping` / `to_u64_wrapping` below).
//!
//! Representation: sign-magnitude, little-endian u64 limbs, canonical
//! form (no trailing zero limbs; zero is `limbs: []`, non-negative).
//! Bitwise operators follow the spec's infinite-width two's-complement
//! semantics via the `~x = -x-1` identities.

use std::fmt;
use std::sync::Arc;

/// Implementation-defined size cap (§6.1.6.2 permits RangeError for
/// oversize results — V8 does the same at ~1G bits). 1<<20 limbs =
/// 64M bits; callers check `would_exceed_cap` and throw RangeError.
pub const MAX_LIMBS: usize = 1 << 20;

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct BigIntVal {
    negative: bool,
    limbs: Vec<u64>,
}

// ── magnitude helpers (little-endian &[u64], canonical) ───────────────

fn mag_norm(mut v: Vec<u64>) -> Vec<u64> {
    while v.last() == Some(&0) {
        v.pop();
    }
    v
}

fn mag_cmp(a: &[u64], b: &[u64]) -> std::cmp::Ordering {
    use std::cmp::Ordering::*;
    if a.len() != b.len() {
        return a.len().cmp(&b.len());
    }
    for i in (0..a.len()).rev() {
        match a[i].cmp(&b[i]) {
            Equal => continue,
            other => return other,
        }
    }
    Equal
}

fn mag_add(a: &[u64], b: &[u64]) -> Vec<u64> {
    let (long, short) = if a.len() >= b.len() { (a, b) } else { (b, a) };
    let mut out = Vec::with_capacity(long.len() + 1);
    let mut carry = 0u64;
    for i in 0..long.len() {
        let (s1, c1) = long[i].overflowing_add(*short.get(i).unwrap_or(&0));
        let (s2, c2) = s1.overflowing_add(carry);
        out.push(s2);
        carry = (c1 as u64) + (c2 as u64);
    }
    if carry > 0 {
        out.push(carry);
    }
    out
}

/// a - b, requires a >= b.
fn mag_sub(a: &[u64], b: &[u64]) -> Vec<u64> {
    let mut out = Vec::with_capacity(a.len());
    let mut borrow = 0u64;
    for i in 0..a.len() {
        let bi = *b.get(i).unwrap_or(&0);
        let (d1, b1) = a[i].overflowing_sub(bi);
        let (d2, b2) = d1.overflowing_sub(borrow);
        out.push(d2);
        borrow = (b1 as u64) + (b2 as u64);
    }
    debug_assert_eq!(borrow, 0, "mag_sub requires a >= b");
    mag_norm(out)
}

fn mag_mul(a: &[u64], b: &[u64]) -> Vec<u64> {
    if a.is_empty() || b.is_empty() {
        return Vec::new();
    }
    let mut out = vec![0u64; a.len() + b.len()];
    for (i, &ai) in a.iter().enumerate() {
        if ai == 0 {
            continue;
        }
        let mut carry = 0u128;
        for (j, &bj) in b.iter().enumerate() {
            let cur = out[i + j] as u128 + (ai as u128) * (bj as u128) + carry;
            out[i + j] = cur as u64;
            carry = cur >> 64;
        }
        let mut k = i + b.len();
        while carry > 0 {
            let cur = out[k] as u128 + carry;
            out[k] = cur as u64;
            carry = cur >> 64;
            k += 1;
        }
    }
    mag_norm(out)
}

fn mag_bit_len(a: &[u64]) -> usize {
    match a.last() {
        None => 0,
        Some(top) => (a.len() - 1) * 64 + (64 - top.leading_zeros() as usize),
    }
}

fn mag_get_bit(a: &[u64], i: usize) -> bool {
    a.get(i / 64)
        .map(|w| (w >> (i % 64)) & 1 == 1)
        .unwrap_or(false)
}

fn mag_set_bit(a: &mut Vec<u64>, i: usize) {
    let limb = i / 64;
    if a.len() <= limb {
        a.resize(limb + 1, 0);
    }
    a[limb] |= 1u64 << (i % 64);
}

fn mag_shl(a: &[u64], bits: usize) -> Vec<u64> {
    if a.is_empty() {
        return Vec::new();
    }
    let limb_shift = bits / 64;
    let bit_shift = bits % 64;
    let mut out = vec![0u64; a.len() + limb_shift + 1];
    for (i, &w) in a.iter().enumerate() {
        out[i + limb_shift] |= w << bit_shift;
        if bit_shift > 0 {
            out[i + limb_shift + 1] |= w >> (64 - bit_shift);
        }
    }
    mag_norm(out)
}

fn mag_shr(a: &[u64], bits: usize) -> Vec<u64> {
    let limb_shift = bits / 64;
    if limb_shift >= a.len() {
        return Vec::new();
    }
    let bit_shift = bits % 64;
    let mut out = Vec::with_capacity(a.len() - limb_shift);
    for i in limb_shift..a.len() {
        let mut w = a[i] >> bit_shift;
        if bit_shift > 0 {
            if let Some(&hi) = a.get(i + 1) {
                w |= hi << (64 - bit_shift);
            }
        }
        out.push(w);
    }
    mag_norm(out)
}

/// Binary long division on magnitudes. b must be non-zero.
/// Returns (quotient, remainder). O(bits·limbs) — correctness first.
fn mag_divrem(a: &[u64], b: &[u64]) -> (Vec<u64>, Vec<u64>) {
    debug_assert!(!b.is_empty());
    if mag_cmp(a, b) == std::cmp::Ordering::Less {
        return (Vec::new(), a.to_vec());
    }
    // Single-limb fast path (also serves parse/to_string).
    if b.len() == 1 {
        let d = b[0] as u128;
        let mut q = vec![0u64; a.len()];
        let mut rem = 0u128;
        for i in (0..a.len()).rev() {
            let cur = (rem << 64) | a[i] as u128;
            q[i] = (cur / d) as u64;
            rem = cur % d;
        }
        let r = if rem == 0 {
            Vec::new()
        } else {
            vec![rem as u64]
        };
        return (mag_norm(q), r);
    }
    let n = mag_bit_len(a);
    let mut q: Vec<u64> = vec![0u64; a.len()];
    let mut r: Vec<u64> = Vec::new();
    for i in (0..n).rev() {
        r = mag_shl(&r, 1);
        if mag_get_bit(a, i) {
            mag_set_bit(&mut r, 0);
        }
        if mag_cmp(&r, b) != std::cmp::Ordering::Less {
            r = mag_sub(&r, b);
            mag_set_bit(&mut q, i);
        }
    }
    (mag_norm(q), r)
}

fn mag_and(a: &[u64], b: &[u64]) -> Vec<u64> {
    let n = a.len().min(b.len());
    mag_norm((0..n).map(|i| a[i] & b[i]).collect())
}

fn mag_or(a: &[u64], b: &[u64]) -> Vec<u64> {
    let n = a.len().max(b.len());
    mag_norm(
        (0..n)
            .map(|i| a.get(i).unwrap_or(&0) | b.get(i).unwrap_or(&0))
            .collect(),
    )
}

fn mag_xor(a: &[u64], b: &[u64]) -> Vec<u64> {
    let n = a.len().max(b.len());
    mag_norm(
        (0..n)
            .map(|i| a.get(i).unwrap_or(&0) ^ b.get(i).unwrap_or(&0))
            .collect(),
    )
}

/// a & !b (b zero-extended).
fn mag_andnot(a: &[u64], b: &[u64]) -> Vec<u64> {
    mag_norm(
        (0..a.len())
            .map(|i| a[i] & !b.get(i).unwrap_or(&0))
            .collect(),
    )
}

impl BigIntVal {
    // ── construction ──────────────────────────────────────────────────

    pub fn zero() -> Self {
        BigIntVal {
            negative: false,
            limbs: Vec::new(),
        }
    }

    fn make(negative: bool, limbs: Vec<u64>) -> Self {
        let limbs = mag_norm(limbs);
        BigIntVal {
            negative: negative && !limbs.is_empty(),
            limbs,
        }
    }

    pub fn from_i64(n: i64) -> Self {
        if n == 0 {
            return Self::zero();
        }
        Self::make(n < 0, vec![n.unsigned_abs()])
    }

    pub fn from_u64(n: u64) -> Self {
        Self::make(false, vec![n])
    }

    pub fn from_i128(n: i128) -> Self {
        if n == 0 {
            return Self::zero();
        }
        let mag = n.unsigned_abs();
        Self::make(n < 0, vec![mag as u64, (mag >> 64) as u64])
    }

    pub fn from_f64(f: f64) -> Self {
        // ToBigInt of an integral Number; caller validates integrality.
        if !f.is_finite() || f == 0.0 {
            return Self::zero();
        }
        let negative = f < 0.0;
        let mut a = f.abs().trunc();
        let mut limbs = Vec::new();
        let base = 18446744073709551616.0f64; // 2^64
        while a >= 1.0 {
            limbs.push((a % base) as u64);
            a = (a / base).trunc();
        }
        Self::make(negative, limbs)
    }

    // ── observers ─────────────────────────────────────────────────────

    pub fn is_zero(&self) -> bool {
        self.limbs.is_empty()
    }

    pub fn is_negative(&self) -> bool {
        self.negative
    }

    pub fn bit_len(&self) -> usize {
        mag_bit_len(&self.limbs)
    }

    pub fn fits_i64(&self) -> bool {
        match self.limbs.len() {
            0 => true,
            1 => self.limbs[0] <= i64::MAX as u64 || (self.negative && self.limbs[0] == 1 << 63),
            _ => false,
        }
    }

    pub fn fits_i32(&self) -> bool {
        self.fits_i64() && {
            let v = self.to_i64_wrapping();
            v >= i32::MIN as i64 && v <= i32::MAX as i64
        }
    }

    pub fn fits_u32(&self) -> bool {
        !self.negative
            && self.limbs.len() <= 1
            && self.limbs.first().copied().unwrap_or(0) <= u32::MAX as u64
    }

    /// js-types JS-API ToBigInt64: the value modulo 2^64, as signed.
    /// The ONLY sanctioned wrap — used exclusively at wasm boundaries.
    pub fn to_i64_wrapping(&self) -> i64 {
        self.to_u64_wrapping() as i64
    }

    /// ToBigUint64: the value modulo 2^64, as unsigned.
    pub fn to_u64_wrapping(&self) -> u64 {
        let low = self.limbs.first().copied().unwrap_or(0);
        if self.negative {
            low.wrapping_neg()
        } else {
            low
        }
    }

    /// §6.1.6.2 mathematical value → Number (f64 rounding applies).
    pub fn to_f64(&self) -> f64 {
        let mut acc = 0.0f64;
        let base = 18446744073709551616.0f64; // 2^64
        for &w in self.limbs.iter().rev() {
            acc = acc * base + w as f64;
        }
        if self.negative { -acc } else { acc }
    }

    /// True when an operation producing roughly `bits` bits must throw
    /// RangeError ("Maximum BigInt size exceeded") — §6.1.6.2 permits
    /// implementation-defined limits surfaced as RangeError.
    pub fn exceeds_cap(bits: usize) -> bool {
        bits / 64 + 1 > MAX_LIMBS
    }

    // ── comparison ────────────────────────────────────────────────────

    pub fn cmp_big(&self, other: &Self) -> std::cmp::Ordering {
        use std::cmp::Ordering::*;
        match (self.negative, other.negative) {
            (false, true) => Greater,
            (true, false) => Less,
            (false, false) => mag_cmp(&self.limbs, &other.limbs),
            (true, true) => mag_cmp(&other.limbs, &self.limbs),
        }
    }

    /// §7.2.13 mixed BigInt/Number comparison on mathematical values.
    /// None when the Number is NaN (all comparisons false).
    pub fn cmp_f64(&self, f: f64) -> Option<std::cmp::Ordering> {
        use std::cmp::Ordering::*;
        if f.is_nan() {
            return None;
        }
        if f == f64::INFINITY {
            return Some(Less);
        }
        if f == f64::NEG_INFINITY {
            return Some(Greater);
        }
        let t = f.trunc();
        let vs_int = self.cmp_big(&Self::from_f64(t));
        if vs_int != Equal {
            return Some(vs_int);
        }
        // Equal integer parts — the fraction decides.
        let frac = f - t;
        Some(if frac > 0.0 {
            Less
        } else if frac < 0.0 {
            Greater
        } else {
            Equal
        })
    }

    // ── arithmetic (§6.1.6.2 — exact) ─────────────────────────────────

    pub fn neg(&self) -> Self {
        Self::make(!self.negative, self.limbs.clone())
    }

    pub fn add(&self, other: &Self) -> Self {
        if self.negative == other.negative {
            Self::make(self.negative, mag_add(&self.limbs, &other.limbs))
        } else {
            match mag_cmp(&self.limbs, &other.limbs) {
                std::cmp::Ordering::Equal => Self::zero(),
                std::cmp::Ordering::Greater => {
                    Self::make(self.negative, mag_sub(&self.limbs, &other.limbs))
                }
                std::cmp::Ordering::Less => {
                    Self::make(other.negative, mag_sub(&other.limbs, &self.limbs))
                }
            }
        }
    }

    pub fn sub(&self, other: &Self) -> Self {
        self.add(&other.neg())
    }

    pub fn mul(&self, other: &Self) -> Self {
        Self::make(
            self.negative != other.negative,
            mag_mul(&self.limbs, &other.limbs),
        )
    }

    /// §6.1.6.2.5/6: quotient truncates toward zero; remainder takes the
    /// dividend's sign. Caller guarantees non-zero divisor.
    pub fn divrem(&self, other: &Self) -> (Self, Self) {
        let (q, r) = mag_divrem(&self.limbs, &other.limbs);
        (
            Self::make(self.negative != other.negative, q),
            Self::make(self.negative, r),
        )
    }

    /// §6.1.6.2.3 exponentiate. Caller rejects negative exponents and
    /// pre-checks the size cap.
    pub fn pow(&self, mut exp: u64) -> Self {
        let mut result = Self::from_i64(1);
        let mut base = self.clone();
        while exp > 0 {
            if exp & 1 == 1 {
                result = result.mul(&base);
            }
            exp >>= 1;
            if exp > 0 {
                base = base.mul(&base);
            }
        }
        result
    }

    pub fn shl(&self, bits: u64) -> Self {
        Self::make(self.negative, mag_shl(&self.limbs, bits as usize))
    }

    /// Arithmetic shift right: floor(x / 2^bits) — negative values round
    /// toward negative infinity (§6.1.6.2.9 via division semantics).
    pub fn shr(&self, bits: u64) -> Self {
        if !self.negative {
            return Self::make(false, mag_shr(&self.limbs, bits as usize));
        }
        // floor for negatives: -((|x| - 1) >> n) - 1
        let m = mag_sub(&self.limbs, &[1]);
        let shifted = mag_shr(&m, bits as usize);
        Self::make(true, mag_add(&shifted, &[1]))
    }

    // ── bitwise (infinite-width two's complement via ~x = -x-1) ───────

    pub fn not(&self) -> Self {
        // ~x = -x - 1
        self.neg().sub(&Self::from_i64(1))
    }

    fn mag_minus_one(&self) -> Vec<u64> {
        // |x| - 1 for negative x (x != 0).
        mag_sub(&self.limbs, &[1])
    }

    pub fn bit_and(&self, other: &Self) -> Self {
        match (self.negative, other.negative) {
            (false, false) => Self::make(false, mag_and(&self.limbs, &other.limbs)),
            (true, true) => {
                // x & y = -(((-x-1) | (-y-1)) + 1)
                let m = mag_or(&self.mag_minus_one(), &other.mag_minus_one());
                Self::make(true, mag_add(&m, &[1]))
            }
            (false, true) => Self::make(false, mag_andnot(&self.limbs, &other.mag_minus_one())),
            (true, false) => Self::make(false, mag_andnot(&other.limbs, &self.mag_minus_one())),
        }
    }

    pub fn bit_or(&self, other: &Self) -> Self {
        match (self.negative, other.negative) {
            (false, false) => Self::make(false, mag_or(&self.limbs, &other.limbs)),
            (true, true) => {
                // x | y = -(((-x-1) & (-y-1)) + 1)
                let m = mag_and(&self.mag_minus_one(), &other.mag_minus_one());
                Self::make(true, mag_add(&m, &[1]))
            }
            (false, true) => {
                // x | y = -(((-y-1) & ~x) + 1)
                let m = mag_andnot(&other.mag_minus_one(), &self.limbs);
                Self::make(true, mag_add(&m, &[1]))
            }
            (true, false) => {
                let m = mag_andnot(&self.mag_minus_one(), &other.limbs);
                Self::make(true, mag_add(&m, &[1]))
            }
        }
    }

    pub fn bit_xor(&self, other: &Self) -> Self {
        match (self.negative, other.negative) {
            (false, false) => Self::make(false, mag_xor(&self.limbs, &other.limbs)),
            (true, true) => Self::make(
                false,
                mag_xor(&self.mag_minus_one(), &other.mag_minus_one()),
            ),
            (false, true) => {
                // x ^ y = -((x ^ (-y-1)) + 1)
                let m = mag_xor(&self.limbs, &other.mag_minus_one());
                Self::make(true, mag_add(&m, &[1]))
            }
            (true, false) => {
                let m = mag_xor(&self.mag_minus_one(), &other.limbs);
                Self::make(true, mag_add(&m, &[1]))
            }
        }
    }

    /// §21.2.2.1 BigInt.asIntN — value modulo 2^bits, sign-extended.
    pub fn as_int_n(&self, bits: u64) -> Self {
        if bits == 0 {
            return Self::zero();
        }
        let m = self.as_uint_n(bits);
        let sign_bit = bits as usize - 1;
        if mag_get_bit(&m.limbs, sign_bit) {
            // m - 2^bits
            let mut p = Vec::new();
            mag_set_bit(&mut p, bits as usize);
            m.sub(&Self::make(false, p))
        } else {
            m
        }
    }

    /// §21.2.2.2 BigInt.asUintN — value modulo 2^bits, unsigned.
    pub fn as_uint_n(&self, bits: u64) -> Self {
        if bits == 0 {
            return Self::zero();
        }
        let full_limbs = (bits as usize).div_ceil(64);
        // Non-negative: mask. Negative: 2^bits - (|x| mod 2^bits).
        let mut masked: Vec<u64> = self.limbs.iter().copied().take(full_limbs).collect();
        if bits % 64 != 0 {
            if let Some(last) = masked.get_mut(full_limbs - 1) {
                *last &= (1u64 << (bits % 64)) - 1;
            }
        }
        let masked = mag_norm(masked);
        if !self.negative || masked.is_empty() {
            return Self::make(false, masked);
        }
        let mut p = Vec::new();
        mag_set_bit(&mut p, bits as usize);
        Self::make(false, mag_sub(&p, &masked))
    }

    // ── string conversion (§7.1.14 / §6.1.6.2.23) ─────────────────────

    /// StringToBigInt body for a known radix. Digits only (sign and
    /// prefixes handled by `parse`).
    fn parse_digits(t: &str, radix: u32) -> Option<Vec<u64>> {
        if t.is_empty() {
            return None;
        }
        let mut mag: Vec<u64> = Vec::new();
        for ch in t.chars() {
            let d = ch.to_digit(radix)? as u64;
            // mag = mag * radix + d
            let mut carry = d as u128;
            for w in mag.iter_mut() {
                let cur = (*w as u128) * (radix as u128) + carry;
                *w = cur as u64;
                carry = cur >> 64;
            }
            while carry > 0 {
                mag.push(carry as u64);
                carry >>= 64;
            }
        }
        Some(mag_norm(mag))
    }

    /// §7.1.14 StringToBigInt: trim, optional sign (decimal only),
    /// 0x/0o/0b prefixes, empty → 0, invalid → None.
    pub fn parse(s: &str) -> Option<Self> {
        let t = s.trim();
        if t.is_empty() {
            return Some(Self::zero());
        }
        if let Some(hex) = t.strip_prefix("0x").or_else(|| t.strip_prefix("0X")) {
            return Self::parse_digits(hex, 16).map(|m| Self::make(false, m));
        }
        if let Some(bin) = t.strip_prefix("0b").or_else(|| t.strip_prefix("0B")) {
            return Self::parse_digits(bin, 2).map(|m| Self::make(false, m));
        }
        if let Some(oct) = t.strip_prefix("0o").or_else(|| t.strip_prefix("0O")) {
            return Self::parse_digits(oct, 8).map(|m| Self::make(false, m));
        }
        let (negative, digits) = match t.strip_prefix('-') {
            Some(rest) => (true, rest),
            None => (false, t.strip_prefix('+').unwrap_or(t)),
        };
        Self::parse_digits(digits, 10).map(|m| Self::make(negative, m))
    }

    /// §6.1.6.2.23 BigInt::toString for radix 2..=36.
    pub fn to_string_radix(&self, radix: u32) -> String {
        if self.is_zero() {
            return "0".to_string();
        }
        let mut mag = self.limbs.clone();
        let mut digits = Vec::new();
        while !mag.is_empty() {
            let (q, r) = mag_divrem(&mag, &[radix as u64]);
            let d = r.first().copied().unwrap_or(0) as u32;
            digits.push(std::char::from_digit(d, radix).unwrap_or('?'));
            mag = q;
        }
        if self.negative {
            digits.push('-');
        }
        digits.iter().rev().collect()
    }
}

impl fmt::Display for BigIntVal {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.to_string_radix(10))
    }
}

impl From<i64> for BigIntVal {
    fn from(n: i64) -> Self {
        Self::from_i64(n)
    }
}

/// Shared-ownership handle — `Value::BigInt` carries this so cloning a
/// Value is a refcount bump, matching String's `Arc<str>` economics.
pub type BigIntRef = Arc<BigIntVal>;
