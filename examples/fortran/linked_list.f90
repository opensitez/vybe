! Linked List and Binary Search Tree
! Covers: derived types with pointer components, recursive procedures,
!         nullify, associated(), pointer assignment, type-bound procedures,
!         generic interfaces, operator overloading.

module linked_list_module
    implicit none
    private

    public :: list_node, linked_list, bst_node, bst

    ! ── Singly-linked list ────────────────────────────────────────────────────

    type :: list_node
        integer :: value
        type(list_node), pointer :: next => null()
    end type list_node

    type :: linked_list
        type(list_node), pointer :: head => null()
        integer :: length = 0
    contains
        procedure :: push_front
        procedure :: push_back
        procedure :: pop_front
        procedure :: contains_value => list_contains
        procedure :: print_list
        procedure :: reverse
        procedure :: to_array
        final     :: list_destructor
    end type linked_list

    ! ── Binary Search Tree ────────────────────────────────────────────────────

    type :: bst_node
        integer :: key
        integer :: count = 1          ! for duplicate keys
        type(bst_node), pointer :: left  => null()
        type(bst_node), pointer :: right => null()
    end type bst_node

    type :: bst
        type(bst_node), pointer :: root => null()
        integer :: size = 0
    contains
        procedure :: insert    => bst_insert
        procedure :: search    => bst_search
        procedure :: delete    => bst_delete
        procedure :: inorder   => bst_inorder
        procedure :: height    => bst_height
        procedure :: min_val   => bst_min
        procedure :: max_val   => bst_max
        final     :: bst_destructor
    end type bst

contains

    ! ── Linked list operations ────────────────────────────────────────────────

    subroutine push_front(self, val)
        class(linked_list), intent(inout) :: self
        integer, intent(in) :: val
        type(list_node), pointer :: node
        allocate(node)
        node%value = val
        node%next  => self%head
        self%head  => node
        self%length = self%length + 1
    end subroutine push_front

    subroutine push_back(self, val)
        class(linked_list), intent(inout) :: self
        integer, intent(in) :: val
        type(list_node), pointer :: node, cur
        allocate(node)
        node%value = val
        nullify(node%next)
        if (.not. associated(self%head)) then
            self%head => node
        else
            cur => self%head
            do while (associated(cur%next))
                cur => cur%next
            end do
            cur%next => node
        end if
        self%length = self%length + 1
    end subroutine push_back

    function pop_front(self) result(val)
        class(linked_list), intent(inout) :: self
        integer :: val
        type(list_node), pointer :: tmp
        if (.not. associated(self%head)) then
            val = -huge(val)
            return
        end if
        val       = self%head%value
        tmp       => self%head
        self%head => self%head%next
        deallocate(tmp)
        self%length = self%length - 1
    end function pop_front

    function list_contains(self, val) result(found)
        class(linked_list), intent(in) :: self
        integer, intent(in) :: val
        logical :: found
        type(list_node), pointer :: cur
        found = .false.
        cur => self%head
        do while (associated(cur))
            if (cur%value == val) then
                found = .true.
                return
            end if
            cur => cur%next
        end do
    end function list_contains

    subroutine print_list(self)
        class(linked_list), intent(in) :: self
        type(list_node), pointer :: cur
        cur => self%head
        write(*, "(a)", advance="no") "["
        do while (associated(cur))
            if (associated(cur%next)) then
                write(*, "(i0, a)", advance="no") cur%value, " -> "
            else
                write(*, "(i0)", advance="no") cur%value
            end if
            cur => cur%next
        end do
        write(*, "(a)") "]"
    end subroutine print_list

    subroutine reverse(self)
        class(linked_list), intent(inout) :: self
        type(list_node), pointer :: prev, cur, nxt
        nullify(prev)
        cur => self%head
        do while (associated(cur))
            nxt      => cur%next
            cur%next => prev
            prev     => cur
            cur      => nxt
        end do
        self%head => prev
    end subroutine reverse

    subroutine to_array(self, arr)
        class(linked_list), intent(in) :: self
        integer, allocatable, intent(out) :: arr(:)
        type(list_node), pointer :: cur
        integer :: i
        allocate(arr(self%length))
        cur => self%head
        i = 1
        do while (associated(cur))
            arr(i) = cur%value
            cur => cur%next
            i = i + 1
        end do
    end subroutine to_array

    subroutine list_destructor(self)
        type(linked_list), intent(inout) :: self
        type(list_node), pointer :: cur, nxt
        cur => self%head
        do while (associated(cur))
            nxt => cur%next
            deallocate(cur)
            cur => nxt
        end do
        nullify(self%head)
    end subroutine list_destructor

    ! ── BST operations ────────────────────────────────────────────────────────

    recursive subroutine bst_insert(self, key)
        class(bst), intent(inout) :: self
        integer, intent(in) :: key
        call insert_node(self%root, key)
        self%size = self%size + 1
    end subroutine bst_insert

    recursive subroutine insert_node(node, key)
        type(bst_node), pointer, intent(inout) :: node
        integer, intent(in) :: key
        if (.not. associated(node)) then
            allocate(node)
            node%key = key
            node%count = 1
            nullify(node%left, node%right)
        else if (key < node%key) then
            call insert_node(node%left, key)
        else if (key > node%key) then
            call insert_node(node%right, key)
        else
            node%count = node%count + 1
        end if
    end subroutine insert_node

    function bst_search(self, key) result(found)
        class(bst), intent(in) :: self
        integer, intent(in) :: key
        logical :: found
        type(bst_node), pointer :: cur
        cur => self%root
        found = .false.
        do while (associated(cur))
            if (key == cur%key) then
                found = .true.
                return
            else if (key < cur%key) then
                cur => cur%left
            else
                cur => cur%right
            end if
        end do
    end function bst_search

    subroutine bst_delete(self, key)
        class(bst), intent(inout) :: self
        integer, intent(in) :: key
        call delete_node(self%root, key)
        self%size = self%size - 1
    end subroutine bst_delete

    recursive subroutine delete_node(node, key)
        type(bst_node), pointer, intent(inout) :: node
        integer, intent(in) :: key
        type(bst_node), pointer :: tmp
        integer :: min_key

        if (.not. associated(node)) return

        if (key < node%key) then
            call delete_node(node%left, key)
        else if (key > node%key) then
            call delete_node(node%right, key)
        else
            ! Found — handle duplicates
            if (node%count > 1) then
                node%count = node%count - 1
                return
            end if
            ! Remove node
            if (.not. associated(node%left)) then
                tmp => node
                node => node%right
                deallocate(tmp)
            else if (.not. associated(node%right)) then
                tmp => node
                node => node%left
                deallocate(tmp)
            else
                ! Two children: replace with in-order successor
                min_key = find_min(node%right)
                node%key = min_key
                call delete_node(node%right, min_key)
            end if
        end if
    end subroutine delete_node

    recursive function find_min(node) result(val)
        type(bst_node), pointer, intent(in) :: node
        integer :: val
        if (associated(node%left)) then
            val = find_min(node%left)
        else
            val = node%key
        end if
    end function find_min

    subroutine bst_inorder(self)
        class(bst), intent(in) :: self
        write(*, "(a)", advance="no") "Inorder: "
        call inorder_traverse(self%root)
        write(*, *)
    end subroutine bst_inorder

    recursive subroutine inorder_traverse(node)
        type(bst_node), pointer, intent(in) :: node
        if (.not. associated(node)) return
        call inorder_traverse(node%left)
        write(*, "(i0, a)", advance="no") node%key, " "
        call inorder_traverse(node%right)
    end subroutine inorder_traverse

    recursive function bst_height(self) result(h)
        class(bst), intent(in) :: self
        integer :: h
        h = node_height(self%root)
    end function bst_height

    recursive function node_height(node) result(h)
        type(bst_node), pointer, intent(in) :: node
        integer :: h, lh, rh
        if (.not. associated(node)) then
            h = 0
        else
            lh = node_height(node%left)
            rh = node_height(node%right)
            h  = 1 + max(lh, rh)
        end if
    end function node_height

    function bst_min(self) result(val)
        class(bst), intent(in) :: self
        integer :: val
        if (associated(self%root)) then
            val = find_min(self%root)
        else
            val = huge(val)
        end if
    end function bst_min

    function bst_max(self) result(val)
        class(bst), intent(in) :: self
        integer :: val
        type(bst_node), pointer :: cur
        if (.not. associated(self%root)) then
            val = -huge(val)
            return
        end if
        cur => self%root
        do while (associated(cur%right))
            cur => cur%right
        end do
        val = cur%key
    end function bst_max

    recursive subroutine free_bst_node(node)
        type(bst_node), pointer, intent(inout) :: node
        if (.not. associated(node)) return
        call free_bst_node(node%left)
        call free_bst_node(node%right)
        deallocate(node)
    end subroutine free_bst_node

    subroutine bst_destructor(self)
        type(bst), intent(inout) :: self
        call free_bst_node(self%root)
    end subroutine bst_destructor

end module linked_list_module


program linked_list_demo
    use linked_list_module
    implicit none

    type(linked_list) :: list
    type(bst)         :: tree
    integer, allocatable :: arr(:)
    integer :: i, val
    integer, parameter :: data(*) = [5, 3, 8, 1, 4, 7, 9, 2, 6, 10, 3, 7]

    ! ── Linked list demo ──────────────────────────────────────────────────────
    print *, "=== Linked List Demo ==="
    print *, ""

    do i = 1, size(data)
        call list%push_back(data(i))
    end do

    write(*, "(a)", advance="no") "Original:  "
    call list%print_list()

    call list%reverse()
    write(*, "(a)", advance="no") "Reversed:  "
    call list%print_list()

    print "(a, i0)", "Length = ", list%length
    print "(a, l1)", "Contains 7? ", list%contains_value(7)
    print "(a, l1)", "Contains 99? ", list%contains_value(99)

    val = list%pop_front()
    print "(a, i0)", "Popped: ", val
    write(*, "(a)", advance="no") "After pop: "
    call list%print_list()

    call list%to_array(arr)
    print "(a, *(i0, 1x))", "As array: ", arr

    ! ── BST demo ──────────────────────────────────────────────────────────────
    print *, ""
    print *, "=== Binary Search Tree Demo ==="
    print *, ""

    do i = 1, size(data)
        call tree%insert(data(i))
    end do

    call tree%inorder()
    print "(a, i0)", "Tree size   = ", tree%size
    print "(a, i0)", "Tree height = ", tree%height()
    print "(a, i0)", "Min value   = ", tree%min_val()
    print "(a, i0)", "Max value   = ", tree%max_val()

    print "(a, l1)", "Search 7?   = ", tree%search(7)
    print "(a, l1)", "Search 42?  = ", tree%search(42)

    print *, ""
    print *, "Deleting 3, 7, 1..."
    call tree%delete(3)
    call tree%delete(7)
    call tree%delete(1)
    call tree%inorder()
    print "(a, i0)", "Tree size after deletions = ", tree%size

    deallocate(arr)

end program linked_list_demo
