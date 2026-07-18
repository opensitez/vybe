use super::helpers::run_prints;

#[test]
fn select_type_polymorphic_matching_type_is_integer() {
    let out = run_prints(
        r#"
program select_type_polymorphic_matching_type_is_integer
    class(*), allocatable :: value
    allocate(integer :: value)
    value = 4
    select type (value)
    type is (integer)
        print *, value
    class default
        print *, -1
    end select
end program select_type_polymorphic_matching_type_is_integer
"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn select_type_polymorphic_matching_class_is_extension_chain() {
    let out = run_prints(
        r#"
program select_type_polymorphic_matching_class_is_extension_chain
    type :: Base
        integer :: a = 1
    end type Base
    type, extends(Base) :: Child
        integer :: b = 5
    end type Child

    class(Base), allocatable :: item
    allocate(Child :: item)
    select type(item)
    class is (Child)
        print *, item%b
    class is (Base)
        print *, item%a
    class default
        print *, -1
    end select
end program select_type_polymorphic_matching_class_is_extension_chain
"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn select_type_polymorphic_matching_logical_dispatch() {
    let out = run_prints(
        r#"
program select_type_polymorphic_matching_logical_dispatch
    class(*), allocatable :: value
    allocate(logical :: value)
    value = .true.
    select type (value)
    type is (logical)
        print *, value
    class default
        print *, .false.
    end select
end program select_type_polymorphic_matching_logical_dispatch
"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn select_type_polymorphic_matching_character_dispatch() {
    let out = run_prints(
        r#"
program select_type_polymorphic_matching_character_dispatch
    class(*), allocatable :: value
    allocate(character(len=4) :: value)
    value = 'data'
    select type (value)
    type is (character(len=*))
        print *, trim(value)
    class default
        print *, 'nope'
    end select
end program select_type_polymorphic_matching_character_dispatch
"#,
    );
    assert_eq!(out, vec!["data"]);
}

#[test]
fn select_type_polymorphic_matching_real_dispatch() {
    let out = run_prints(
        r#"
program select_type_polymorphic_matching_real_dispatch
    class(*), allocatable :: value
    allocate(real :: value)
    value = 3.5
    select type (value)
    type is (real)
        print *, int(value)
    class default
        print *, 0
    end select
end program select_type_polymorphic_matching_real_dispatch
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn select_type_polymorphic_matching_arrayed_polymorph() {
    let out = run_prints(
        r#"
program select_type_polymorphic_matching_arrayed_polymorph
    class(*), allocatable :: value
    allocate(integer :: value(3))
    value = (/1, 2, 3/)
    select type (value)
    type is (integer)
        print *, value(1)
    class default
        print *, -1
    end select
end program select_type_polymorphic_matching_arrayed_polymorph
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn select_type_polymorphic_matching_nested_selector_fallthrough() {
    let out = run_prints(
        r#"
program select_type_polymorphic_matching_nested_selector_fallthrough
    class(*), allocatable :: container
    allocate(integer :: container)
    container = 9
    select type (container)
    type is (real)
        print *, 1
    class is (integer)
        print *, 2
    class default
        print *, 3
    end select
end program select_type_polymorphic_matching_nested_selector_fallthrough
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn select_type_polymorphic_matching_default_only() {
    let out = run_prints(
        r#"
program select_type_polymorphic_matching_default_only
    class(*), allocatable :: marker
    allocate(integer(2) :: marker)
    select type (marker)
    class is (logical)
        print *, 0
    class default
        print *, size(marker)
    end select
end program select_type_polymorphic_matching_default_only
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn select_type_polymorphic_matching_derived_component_polymorph() {
    let out = run_prints(
        r#"
program select_type_polymorphic_matching_derived_component_polymorph
    type :: Packet
        integer :: n = 2
    end type Packet
    class(*) :: holder
    class(Packet), allocatable :: payload
    allocate(Packet :: payload)
    holder = payload%n
    select type (payload)
    type is (Packet)
        print *, payload%n
    class default
        print *, -1
    end select
end program select_type_polymorphic_matching_derived_component_polymorph
"#,
    );
    assert_eq!(out, vec!["2"]);
}
