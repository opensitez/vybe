pub mod fixtures {
    include!(concat!(env!("OUT_DIR"), "/fixtures.rs"));
}
pub mod islands {
    include!(concat!(env!("OUT_DIR"), "/islands.rs"));
}
pub mod lua {
    include!(concat!(env!("OUT_DIR"), "/lua.rs"));
}
pub mod lua_pratt {
    include!(concat!(env!("OUT_DIR"), "/lua_pratt.rs"));
}
pub mod ast {
    include!(concat!(env!("OUT_DIR"), "/ast.rs"));
}
pub mod expression {
    include!(concat!(env!("OUT_DIR"), "/expression.rs"));
}
pub mod modular {
    include!(concat!(env!("OUT_DIR"), "/modular.rs"));
}
#[cfg(feature = "all-grammars")]
pub mod all {
    include!(concat!(env!("OUT_DIR"), "/all_grammars.rs"));
}
