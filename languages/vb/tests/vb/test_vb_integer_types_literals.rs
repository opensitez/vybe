use super::helpers::run_vb;

#[test] fn int_literal_short() { assert_eq!(run_vb("Module M\nSub Main()\nDim x As Short = 32767\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Int16"]); }
#[test] fn int_literal_short_type_char() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 10S\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Int16"]); }
#[test] fn int_literal_integer() { assert_eq!(run_vb("Module M\nSub Main()\nDim x As Integer = 2147483647\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Int32"]); }
#[test] fn int_literal_integer_type_char() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 10I\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Int32"]); }
#[test] fn int_literal_integer_type_char_legacy() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 10%\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Int32"]); }

#[test] fn int_literal_long() { assert_eq!(run_vb("Module M\nSub Main()\nDim x As Long = 9223372036854775807\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Int64"]); }
#[test] fn int_literal_long_type_char() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 10L\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Int64"]); }
#[test] fn int_literal_long_type_char_legacy() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 10&\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Int64"]); }

#[test] fn int_literal_byte() { assert_eq!(run_vb("Module M\nSub Main()\nDim x As Byte = 255\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Byte"]); }
#[test] fn int_literal_sbyte() { assert_eq!(run_vb("Module M\nSub Main()\nDim x As SByte = -128\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"), vec!["SByte"]); }
#[test] fn int_literal_ushort() { assert_eq!(run_vb("Module M\nSub Main()\nDim x As UShort = 65535\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"), vec!["UInt16"]); }
#[test] fn int_literal_ushort_type_char() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 10US\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"), vec!["UInt16"]); }

#[test] fn int_literal_uinteger() { assert_eq!(run_vb("Module M\nSub Main()\nDim x As UInteger = 4294967295\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"), vec!["UInt32"]); }
#[test] fn int_literal_uinteger_type_char() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 10UI\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"), vec!["UInt32"]); }

#[test] fn int_literal_ulong() { assert_eq!(run_vb("Module M\nSub Main()\nDim x As ULong = 18446744073709551615UL\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"), vec!["UInt64"]); }
#[test] fn int_literal_ulong_type_char() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 10UL\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"), vec!["UInt64"]); }

#[test] fn int_literal_hex() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = &HFF\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["255"]); }
#[test] fn int_literal_hex_long() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = &HFFFFFFFFL\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Int64"]); }
#[test] fn int_literal_octal() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = &O10\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["8"]); }
#[test] fn int_literal_octal_long() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = &O10L\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Int64"]); }
#[test] fn int_literal_binary() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = &B1010\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["10"]); }
#[test] fn int_literal_binary_long() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = &B1010L\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Int64"]); }
#[test] fn int_literal_hex_unsigned() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = &HFFFFFFFFUI\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["4294967295"]); }
#[test] fn int_literal_overflow_to_long() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 2147483648\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Int64"]); }
#[test] fn int_literal_underscore_separator() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 1_000_000\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["1000000"]); }
