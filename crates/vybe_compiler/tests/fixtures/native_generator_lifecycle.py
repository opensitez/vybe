def producer(values):
    for value in values:
        yield value

def generator_min(values):
    return min(producer(values))

def captured_min(values):
    offset = 10
    def captured():
        for value in values:
            yield value + offset
    return min(captured())

def generator_make(values):
    return producer(values)

def generator_next(iterator):
    return next(iterator)

def generator_close(iterator):
    iterator.close()
    return 17

def generator_throw(iterator):
    return iterator.throw(ValueError("boom"))

starts = 0

def counted_producer():
    global starts
    starts += 1
    yield 7

def fresh_close_count():
    global starts
    starts = 0
    iterator = counted_producer()
    iterator.close()
    return starts

def generator_done_status(iterator):
    try:
        next(iterator)
    except StopIteration:
        return "done"
    return "value"

def source_throw_done_status(values):
    iterator = generator_make(values)
    next(iterator)
    try:
        generator_throw(iterator)
    except ValueError:
        try:
            next(iterator)
        except StopIteration:
            return "done"
    return "not-done"
