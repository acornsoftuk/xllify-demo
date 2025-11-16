import xllify
import numpy as np

# start 3
xllify.configure_spawn_count(3)

@xllify.fn("test_func_py")
def test_func(x: any):
    return np.random.random()

@xllify.fn("test_func_range")
def test_func(x: int):
    return np.random.random(x)
