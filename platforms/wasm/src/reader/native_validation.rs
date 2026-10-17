//! Standard validation; extra syntax diagnostics run only for rejected modules.

use super::WasmError;
use crate::binary_cursor::Cursor;
use crate::binary_sections::{BinaryKind, Sections};
use wasmparser::{Validator, WasmFeatures};

fn check_function_syntax(data: &[u8]) -> Result<(), WasmError> {
    let sections = Sections::new(data)?;
    if sections.kind() != BinaryKind::Module {
        return Err("expected a core Wasm module".into());
    }
    for section in sections {
        let section = section?;
        if section.id != 10 {
            continue;
        }
        let mut code = Cursor::new(section.payload);
        let count = super::section_types::count(&mut code, 1)?;
        for _ in 0..count {
            let size = code.u32()? as usize;
            let mut body = Cursor::new(code.bytes(size)?);
            let groups = super::section_types::count(&mut body, 2)?;
            for _ in 0..groups {
                body.u32()?;
                super::section_types::value_type(&mut body)?;
            }
            while body.remaining() != 0 {
                crate::binary_instructions::next(&mut body)?;
            }
        }
        super::section_types::finish(&code)?;
    }
    Ok(())
}

pub(crate) fn validate_component(data: &[u8]) -> Result<(), WasmError> {
    Validator::new_with_features(WasmFeatures::all())
        .validate_all(data)
        .map(|_| ())
        .map_err(|error| WasmError::invalid(format!("invalid component binary: {error}")))
}

pub(super) fn validate(data: &[u8]) -> Result<(), WasmError> {
    match super::typed_core::validate(data) {
        Ok(super::typed_core::Coverage::Complete) => return Ok(()),
        Ok(super::typed_core::Coverage::RequiresProposalValidation) => {}
        Err(error) => {
            check_function_syntax(data)?;
            return Err(error);
        }
    }
    let features = WasmFeatures::all();
    let mut validator = Validator::new_with_features(features);
    let Err(error) = validator.validate_all(data) else {
        return Ok(());
    };

    // The standard validator mixes binary-reader and typing errors. Inspect
    // function syntax only after rejection to retain the reader's phase API.
    check_function_syntax(data)?;
    Err(WasmError::invalid(error.to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::encoding::*;

    fn function_module(result: u8, body: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::from(WASM_MAGIC);
        bytes.extend_from_slice(&WASM_VERSION);
        write_section(&mut bytes, SECTION_TYPE, &[1, TYPE_FUNC, 0, 1, result]);
        write_section(&mut bytes, SECTION_FUNCTION, &[1, 0]);
        let mut code = vec![1];
        write_leb128_u32(&mut code, body.len() as u32);
        code.extend_from_slice(body);
        write_section(&mut bytes, SECTION_CODE, &code);
        bytes
    }

    #[test]
    fn typed_function_results_are_validated() {
        validate(&function_module(TYPE_I32, &[0, 0x41, 42, 0x0b])).unwrap();
        let missing = validate(&function_module(TYPE_I32, &[0, 0x0b])).unwrap_err();
        assert_eq!(missing.phase, super::super::Phase::Invalid);
        let wrong = validate(&function_module(TYPE_I32, &[0, 0x42, 42, 0x0b])).unwrap_err();
        assert_eq!(wrong.phase, super::super::Phase::Invalid);
    }

    #[test]
    fn malformed_operators_remain_decode_errors() {
        let error = validate(&function_module(TYPE_I32, &[0, 0xff, 0x0b])).unwrap_err();
        assert_eq!(error.phase, super::super::Phase::Malformed);
    }

    #[test]
    fn inhouse_syntax_diagnostics_distinguish_immediates_from_typing() {
        // Bytes resembling opcodes inside an f32 immediate are opaque.
        check_function_syntax(&function_module(
            TYPE_I32,
            &[0, 0x43, 0xff, 0xff, 0xff, 0xff, 0x0b],
        ))
        .unwrap();
        for body in [
            &[0, 0x41, 0x80][..],
            &[0, 0x43, 0, 0, 0],
            &[1, 1, 0x63],
            &[0, 0xfb, 0xff, 0x7f],
        ] {
            assert_eq!(
                check_function_syntax(&function_module(TYPE_I32, body))
                    .unwrap_err()
                    .phase,
                super::super::Phase::Malformed
            );
        }
        let wrong_type = function_module(TYPE_I32, &[0, 0x42, 42, 0x0b]);
        check_function_syntax(&wrong_type).unwrap();
        assert_eq!(
            validate(&wrong_type).unwrap_err().phase,
            super::super::Phase::Invalid
        );
    }

    #[test]
    fn invalid_gc_global_initializers_are_rejected_before_instantiation() {
        for (types, globals) in [
            // A struct's i32 field cannot be initialized from i64.
            (
                &[1, 0x5f, 1, TYPE_I32, 0][..],
                &[1, 0x64, 0, 0, 0x42, 42, 0xfb, 0, 0, 0x0b][..],
            ),
            // A non-null function field has no default value.
            (
                &[2, TYPE_FUNC, 0, 0, 0x5f, 1, 0x64, 0, 0][..],
                &[1, 0x64, 1, 0, 0xfb, 1, 1, 0x0b][..],
            ),
        ] {
            let mut bytes = Vec::from(WASM_MAGIC);
            bytes.extend_from_slice(&WASM_VERSION);
            write_section(&mut bytes, SECTION_TYPE, types);
            write_section(&mut bytes, SECTION_GLOBAL, globals);
            let error = super::super::read_wasm_classified(&bytes).unwrap_err();
            assert_eq!(error.phase, super::super::Phase::Invalid);
        }
    }

    #[test]
    fn custom_metadata_does_not_bypass_standard_validation() {
        let mut bytes = function_module(TYPE_I32, &[0, 0x0b]);
        let mut custom = Vec::new();
        write_name(&mut custom, "vybe");
        write_section(&mut bytes, SECTION_CUSTOM, &custom);
        let error = super::super::read_wasm_classified(&bytes).unwrap_err();
        assert_eq!(error.phase, super::super::Phase::Invalid);
    }
}
