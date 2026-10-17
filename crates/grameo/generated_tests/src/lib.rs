pub mod fixtures {
    include!(concat!(env!("OUT_DIR"), "/fixtures.rs"));
}
pub mod capture_api {
    include!(concat!(env!("OUT_DIR"), "/capture_api.rs"));
}
pub mod hybrid_api {
    include!(concat!(env!("OUT_DIR"), "/hybrid_api.rs"));
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

pub mod terminal_index {
    include!(concat!(env!("OUT_DIR"), "/terminal_index.rs"));
}

pub mod quoted {
    include!(concat!(env!("OUT_DIR"), "/quoted.rs"));
}

pub mod php {
    include!(concat!(env!("OUT_DIR"), "/php.rs"));
}
pub mod php_baseline {
    include!(concat!(env!("OUT_DIR"), "/php_baseline.rs"));
}
pub mod php_semantic_baseline {
    include!(concat!(env!("OUT_DIR"), "/php_semantic_baseline.rs"));
}
pub mod php_delimited_baseline {
    include!(concat!(env!("OUT_DIR"), "/php_delimited_baseline.rs"));
}
