import time
from threading import Thread
import functools


class timeclass():
    def __init__(self, testString):
        self.testString = testString

    def timeout(seconds_before_timeout):
        def deco(func):
            @functools.wraps(func)
            def wrapper(*args, **kwargs):
                res = [
                    Exception('function [%s] timeout [%s seconds] exceeded!' % (func.__name__, seconds_before_timeout))]

                def newFunc():
                    try:
                        res[0] = func(*args, **kwargs)
                    except Exception as e:
                        res[0] = e

                t = Thread(target=newFunc)
                t.daemon = True
                try:
                    t.start()
                    t.join(seconds_before_timeout)
                except Exception as e:
                    print('error starting thread')
                    raise e
                ret = res[0]
                if isinstance(ret, BaseException):
                    raise ret
                return ret

            return wrapper

        return deco

    @timeout(12)
    def mytest(self):
        print("Start")
        for i in range(1, 10):
            time.sleep(1)
            print("{} seconds have passed".format(i))
        return self.testString
