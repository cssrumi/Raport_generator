from time import clock

def print_generator(func):
    def wrapper(*args, **kwargs):
        for rv in func(*args, **kwargs):
            print(rv)

    return wrapper


def new_timer(func):
    def wrapper(*args, **kwargs):
        before = clock()
        rv = func(*args, **kwargs)
        after = clock()
        # print(func.__name__, 'elapsed', after - before)
        print('%s elapsed %.7f' % (func.__name__, after - before))
        return rv

    return wrapper
