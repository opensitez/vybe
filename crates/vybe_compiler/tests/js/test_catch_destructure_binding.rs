//! Catch destructure bindings — object, array, nested, defaults, optional catch.

crate::js_cases! {
    catch_destructure_object_message_from_error => {
        r#"try{throw new Error("failed");}catch({message}){console.log(message);}"#,
        ["failed"]
    };

    catch_destructure_object_code_and_message => {
        r#"try{throw{code:404,message:"nf"};}catch({code,message}){console.log(code);console.log(message);}"#,
        ["404", "nf"]
    };

    catch_destructure_array_pair => {
        r#"try{throw["alpha","beta"];}catch([first,second]){console.log(first);console.log(second);}"#,
        ["alpha", "beta"]
    };

    catch_destructure_array_with_rest => {
        r#"try{throw[1,2,3,4];}catch([a,b,...rest]){console.log(a);console.log(rest.length);}"#,
        ["1", "2"]
    };

    catch_destructure_nested_object => {
        r#"try{throw{err:{code:"E"}};}catch({err:{code}}){console.log(code);}"#,
        ["E"]
    };

    catch_destructure_default_when_property_missing => {
        r#"try{throw{};}catch({message="fallback"}){console.log(message);}"#,
        ["fallback"]
    };

    catch_destructure_rename_with_colon => {
        r#"try{throw{status:500};}catch({status:code}){console.log(code);}"#,
        ["500"]
    };

    catch_destructure_array_skip_with_empty_slot => {
        r#"try{throw[10,,30];}catch([a,,c]){console.log(a);console.log(c);}"#,
        ["10", "30"]
    };

    catch_optional_binding_without_using_error => {
        r#"try{throw new Error("x");}catch{console.log("handled");}"#,
        ["handled"]
    };

    catch_destructure_on_thrown_primitive_string => {
        r#"try{throw"plain";}catch(msg){console.log(msg);}"#,
        ["plain"]
    };

    catch_destructure_object_with_symbol_key_via_computed => {
        r#"const k=Symbol("k");try{throw{[k]:9};}catch(e){console.log(e[k]);}"#,
        ["9"]
    };

    catch_destructure_nested_array_in_object => {
        r#"try{throw{pair:[3,4]};}catch({pair:[x,y]}){console.log(x+y);}"#,
        ["7"]
    };

    catch_destructure_reassign_extracted_binding => {
        r#"try{throw new Error("orig");}catch({message}){message="new";console.log(message);}"#,
        ["new"]
    };

    catch_destructure_null_throw_binding => {
        r#"try{throw null;}catch(e){console.log(e===null);}"#,
        ["true"]
    };

    catch_destructure_undefined_throw_binding => {
        r#"try{throw undefined;}catch(e){console.log(e===undefined);}"#,
        ["true"]
    };

    catch_destructure_inside_nested_try => {
        r#"let o=[];try{throw{layer:1};}catch({layer}){try{throw{layer:layer+1};}catch({layer:l}){o.push(l);}}console.log(o.join(","));"#,
        ["2"]
    };

    catch_destructure_array_single_element => {
        r#"try{throw[42];}catch([only]){console.log(only);}"#,
        ["42"]
    };

    catch_destructure_object_getter_throw => {
        r#"try{throw{get msg(){throw new Error("getter");}};}catch({msg}){console.log("skip");}catch(e){console.log(e.message);}"#,
        ["getter"]
    };
}
