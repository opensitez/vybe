use vybe_runtime::chunk::{LocalName, LocalOrigin};
/// Compiler-only offset metadata. The caller supplies its already-checked
/// instruction splice map; source names never activate native lowering.
pub(crate) fn relocate(locals: &mut [LocalName], mut remap: impl FnMut(usize) -> usize) {
    for local in locals {
        if local.origin == LocalOrigin::Compiler {
            if let Some(value) = local.name.strip_prefix("<native-growable-array:").and_then(|s|s.strip_suffix('>')) {
                let offset:usize=value.parse().expect("native growable allocation offset");
                local.name=format!("<native-growable-array:{}>",remap(offset));
                continue;
            }
            if let Some(fields) = local
                .name
                .strip_prefix("<native-generator-resume:")
                .and_then(|s| s.strip_suffix('>'))
            {
                let fields: Vec<usize> = fields
                    .split(':')
                    .map(|x| x.parse().expect("native resume marker offset"))
                    .collect();
                assert_eq!(fields.len(), 2);
                local.name = format!(
                    "<native-generator-resume:{}:{}>",
                    remap(fields[0]),
                    remap(fields[1])
                );
                continue;
            }
        }
        if local.origin != LocalOrigin::Compiler {
            continue;
        }
        if let Some(fields) = local
            .name
            .strip_prefix("<native-generator-next:")
            .and_then(|s| s.strip_suffix('>'))
        {
            let fields: Vec<usize> = fields
                .split(':')
                .map(|x| x.parse().expect("native next marker offset"))
                .collect();
            assert_eq!(fields.len(), 4);
            let moved: Vec<_> = fields.into_iter().map(&mut remap).collect();
            local.name = format!(
                "<native-generator-next:{}:{}:{}:{}>",
                moved[0], moved[1], moved[2], moved[3]
            );
        }
    }
}
