import numpy as np
import matplotlib.pyplot as plt
from typing import *

def generate_random_data(num_points: int, n_randomization: int=3):
    x = np.random.randn(num_points)

    for _ in range(n_randomization):
        random_slice = int(np.random.rand() * num_points)
        x[random_slice:] += np.random.randint(0, 10)

    return x

plt.hist(generate_random_data(5000), bins=100)
plt.show()


def kde(bandwith: float, data: List, kernel: Callable):
    mixture = np.zeros(1000)
    points = np.linspace(0, max(data), 1000)
    for xi in data:
        mixture += kernel(points, xi, bandwith)

    return mixture

def gaussian(x: Any, xi: float, bandwith: float):
    exp_section = np.exp(-np.power(x - xi, 2.0))
    return exp_section / (2 * np.power(bandwith, 2.0))

dist = generate_random_data(500)
hist = np.histogram(dist, bins=50)[1]

plt.figure(
    figsize=(16, 10)
)
plt.hist(dist, bins=50, 
         alpha=0.5, color='green', label='True Distribution')
plt.show()


points = kde(0.5, dist, gaussian)

plt.fill_between(np.linspace(0, max(dist), 1000), 
                 points, alpha=0.5, label='Estimated PDF')

plt.legend(loc='upper right')
plt.show()

points = kde(0.5, dist, gaussian)
points /= np.abs(points).max(axis=0)  # Normalizing
