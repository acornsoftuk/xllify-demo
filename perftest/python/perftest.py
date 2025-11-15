import xllify
import time
import numpy as np

@xllify.fn("test_func")
def test_func(x: any):
    return np.random.random()
