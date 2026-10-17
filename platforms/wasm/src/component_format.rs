//! Shared wire facts used by the component reader and writer.
macro_rules! primitive_types {
    ($($byte:literal => $name:literal),* $(,)?) => {
        pub(crate) fn primitive_name(byte:u8)->Option<&'static str>{
            match byte {$($byte=>Some($name),)* _=>None}
        }
        pub(crate) fn primitive_byte(name:&str)->Option<u8>{
            match name {$($name=>Some($byte),)* _=>None}
        }
    }
}
primitive_types! {
    0x7f=>"bool",0x7e=>"s8",0x7d=>"u8",0x7c=>"s16",0x7b=>"u16",
    0x7a=>"s32",0x79=>"u32",0x78=>"s64",0x77=>"u64",0x76=>"f32",
    0x75=>"f64",0x74=>"char",0x73=>"string",0x64=>"error-context",
}
pub(crate) const HEADER: &[u8; 8] = b"\0asm\x0d\0\x01\0";

#[cfg(test)]
mod tests {
    #[test]
    fn primitive_wire_mapping_is_bijective() {
        let mut count = 0;
        for byte in 0..=u8::MAX {
            if let Some(name) = super::primitive_name(byte) {
                assert_eq!(super::primitive_byte(name), Some(byte));
                count += 1;
            }
        }
        assert_eq!(count, 14);
    }
}
