use vybec::parser_python::ast::*;

fn parse(src: &str) -> Module {
    vybec::parser_python::parse(src).expect("parse failed")
}

#[test]
fn import_simple() {
    let m = parse("import os\n");
    match &m.body[0] {
        Statement::Import { names } => {
            assert_eq!(names.len(), 1);
            assert_eq!(names[0].name, "os");
            assert!(names[0].asname.is_none());
        }
        other => panic!("expected Import, got: {:?}", other),
    }
}

#[test]
fn import_multiple() {
    let m = parse("import os, sys\n");
    match &m.body[0] {
        Statement::Import { names } => {
            assert_eq!(names.len(), 2);
        }
        other => panic!("expected Import, got: {:?}", other),
    }
}

#[test]
fn import_as() {
    let m = parse("import numpy as np\n");
    match &m.body[0] {
        Statement::Import { names } => {
            assert_eq!(names[0].name, "numpy");
            assert_eq!(names[0].asname.as_deref(), Some("np"));
        }
        other => panic!("expected Import, got: {:?}", other),
    }
}

#[test]
fn import_dotted() {
    let m = parse("import os.path\n");
    match &m.body[0] {
        Statement::Import { names } => {
            assert_eq!(names[0].name, "os.path");
        }
        other => panic!("expected Import, got: {:?}", other),
    }
}

#[test]
fn from_import() {
    let m = parse("from os import path\n");
    match &m.body[0] {
        Statement::ImportFrom { module, names, level } => {
            assert_eq!(module.as_deref(), Some("os"));
            assert_eq!(names.len(), 1);
            assert_eq!(names[0].name, "path");
            assert_eq!(*level, 0);
        }
        other => panic!("expected ImportFrom, got: {:?}", other),
    }
}

#[test]
fn from_import_multiple() {
    let m = parse("from pathlib import Path, PurePath\n");
    match &m.body[0] {
        Statement::ImportFrom { names, .. } => {
            assert_eq!(names.len(), 2);
        }
        other => panic!("expected ImportFrom, got: {:?}", other),
    }
}

#[test]
fn from_import_star() {
    let m = parse("from os import *\n");
    match &m.body[0] {
        Statement::ImportFrom { names, .. } => {
            assert_eq!(names.len(), 1);
            assert_eq!(names[0].name, "*");
        }
        other => panic!("expected ImportFrom, got: {:?}", other),
    }
}

#[test]
fn from_import_relative() {
    let m = parse("from . import utils\n");
    match &m.body[0] {
        Statement::ImportFrom { level, module, .. } => {
            assert_eq!(*level, 1);
            assert!(module.is_none());
        }
        other => panic!("expected ImportFrom, got: {:?}", other),
    }
}

#[test]
fn from_import_relative_dotted() {
    let m = parse("from ..models import User\n");
    match &m.body[0] {
        Statement::ImportFrom { level, module, .. } => {
            assert_eq!(*level, 2);
            assert_eq!(module.as_deref(), Some("models"));
        }
        other => panic!("expected ImportFrom, got: {:?}", other),
    }
}

#[test]
fn from_import_with_parens() {
    let m = parse("from os import (\n    path,\n    getcwd\n)\n");
    match &m.body[0] {
        Statement::ImportFrom { names, .. } => {
            assert_eq!(names.len(), 2);
        }
        other => panic!("expected ImportFrom, got: {:?}", other),
    }
}
