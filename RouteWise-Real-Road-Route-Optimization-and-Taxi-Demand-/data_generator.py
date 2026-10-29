"""Deterministic synthetic mobility data for the RouteWise academic demo.

This module models travel demand between 12 regional landmarks in the
Karunagappally study area using a spatial interaction Gravity Model:
    T_ij = k * (P_i^alpha * A_j^beta) / (d_ij^gamma)
where:
  - P_i is the origin landmark trip-generation (population/transit) weight
  - A_j is the destination landmark trip-attraction weight
  - d_ij is the real-road network distance (km) between landmarks
  - gamma = 1.3 is the calibrated distance-decay exponent
  - k is a scaling constant such that total trips N = 5,000.

Temporal sampling incorporates realistic diurnal fluctuations with peak
hours (08:00-10:00 morning office/student commute and 16:30-19:30 evening return)
and off-peak nighttime travel.

Fare formula follows the academic benchmark grounded in Kerala Motor Vehicles
Department (MVD) auto/taxi fare structures:
    Fare = FARE_BASE + (distance_km * FARE_PER_KM) + (travel_time_min * FARE_PER_MIN)
"""
from __future__ import annotations

import networkx as nx
import numpy as np
import pandas as pd

SEED = 42
TRIP_COUNT = 5_000
FARE_BASE = 45.0          # Base flag drop in INR (includes first 1.5 km equivalent)
FARE_PER_KM = 18.0        # Rate per kilometer in INR
FARE_PER_MIN = 1.5        # Congestion / waiting / transit time rate in INR per minute

LOCATIONS = [
    "Karunagappally", "Railway Station", "KSRTC Bus Stand", "Oachira",
    "Chavara", "Sasthamkotta", "Mynagappally", "Thazhava",
    "Clappana", "Vallikavu", "Amritapuri", "Puthiyakavu"
]

# Layout coordinates for 2D diagrammatic / topological visualizations
LAYOUT = {
    name: point for name, point in zip(
        LOCATIONS,
        [(49, 48), (40, 29), (53, 28), (56, 76), (79, 73), (82, 29),
         (67, 52), (26, 31), (18, 53), (29, 72), (17, 78), (71, 16)]
    )
}

# Real-road shortest-path distances (km) computed on the cached OSM MultiDiGraph
EXACT_ROAD_DISTANCES = {
    "Karunagappally": {
        "Karunagappally": 0.0, "Railway Station": 0.89, "KSRTC Bus Stand": 0.16,
        "Oachira": 9.36, "Chavara": 8.24, "Sasthamkotta": 11.64, "Mynagappally": 5.77,
        "Thazhava": 5.25, "Clappana": 8.29, "Vallikavu": 7.77, "Amritapuri": 9.05, "Puthiyakavu": 2.90
    },
    "Railway Station": {
        "Karunagappally": 1.08, "Railway Station": 0.0, "KSRTC Bus Stand": 1.08,
        "Oachira": 10.44, "Chavara": 7.86, "Sasthamkotta": 11.96, "Mynagappally": 5.40,
        "Thazhava": 6.32, "Clappana": 9.32, "Vallikavu": 8.76, "Amritapuri": 10.05, "Puthiyakavu": 3.22
    },
    "KSRTC Bus Stand": {
        "Karunagappally": 0.68, "Railway Station": 0.73, "KSRTC Bus Stand": 0.0,
        "Oachira": 10.01, "Chavara": 8.08, "Sasthamkotta": 11.48, "Mynagappally": 5.61,
        "Thazhava": 5.84, "Clappana": 8.97, "Vallikavu": 8.45, "Amritapuri": 9.74, "Puthiyakavu": 2.74
    },
    "Oachira": {
        "Karunagappally": 9.92, "Railway Station": 9.97, "KSRTC Bus Stand": 9.23,
        "Oachira": 0.0, "Chavara": 17.31, "Sasthamkotta": 20.19, "Mynagappally": 14.85,
        "Thazhava": 4.90, "Clappana": 5.26, "Vallikavu": 9.95, "Amritapuri": 8.78, "Puthiyakavu": 9.22
    },
    "Chavara": {
        "Karunagappally": 7.94, "Railway Station": 7.86, "KSRTC Bus Stand": 8.10,
        "Oachira": 17.30, "Chavara": 0.0, "Sasthamkotta": 13.40, "Mynagappally": 4.24,
        "Thazhava": 13.19, "Clappana": 16.18, "Vallikavu": 15.44, "Amritapuri": 16.73, "Puthiyakavu": 10.84
    },
    "Sasthamkotta": {
        "Karunagappally": 12.17, "Railway Station": 11.96, "KSRTC Bus Stand": 11.48,
        "Oachira": 20.18, "Chavara": 13.40, "Sasthamkotta": 0.0, "Mynagappally": 9.82,
        "Thazhava": 15.71, "Clappana": 19.38, "Vallikavu": 19.79, "Amritapuri": 21.07, "Puthiyakavu": 11.17
    },
    "Mynagappally": {
        "Karunagappally": 5.47, "Railway Station": 5.40, "KSRTC Bus Stand": 5.63,
        "Oachira": 14.83, "Chavara": 4.24, "Sasthamkotta": 9.82, "Mynagappally": 0.0,
        "Thazhava": 10.72, "Clappana": 13.71, "Vallikavu": 12.98, "Amritapuri": 14.26, "Puthiyakavu": 7.75
    },
    "Thazhava": {
        "Karunagappally": 5.79, "Railway Station": 5.84, "KSRTC Bus Stand": 5.11,
        "Oachira": 4.91, "Chavara": 13.19, "Sasthamkotta": 15.71, "Mynagappally": 10.72,
        "Thazhava": 0.0, "Clappana": 4.11, "Vallikavu": 12.11, "Amritapuri": 11.10, "Puthiyakavu": 4.73
    },
    "Clappana": {
        "Karunagappally": 8.29, "Railway Station": 9.01, "KSRTC Bus Stand": 8.27,
        "Oachira": 5.26, "Chavara": 16.19, "Sasthamkotta": 19.38, "Mynagappally": 13.72,
        "Thazhava": 4.11, "Clappana": 0.0, "Vallikavu": 8.84, "Amritapuri": 7.67, "Puthiyakavu": 8.41
    },
    "Vallikavu": {
        "Karunagappally": 7.77, "Railway Station": 8.49, "KSRTC Bus Stand": 7.75,
        "Oachira": 9.95, "Chavara": 15.44, "Sasthamkotta": 19.24, "Mynagappally": 12.98,
        "Thazhava": 12.11, "Clappana": 8.84, "Vallikavu": 0.0, "Amritapuri": 1.29, "Puthiyakavu": 10.31
    },
    "Amritapuri": {
        "Karunagappally": 9.05, "Railway Station": 9.77, "KSRTC Bus Stand": 9.04,
        "Oachira": 8.78, "Chavara": 16.73, "Sasthamkotta": 20.52, "Mynagappally": 14.26,
        "Thazhava": 11.12, "Clappana": 7.67, "Vallikavu": 1.29, "Amritapuri": 0.0, "Puthiyakavu": 11.59
    },
    "Puthiyakavu": {
        "Karunagappally": 3.43, "Railway Station": 3.22, "KSRTC Bus Stand": 2.74,
        "Oachira": 9.21, "Chavara": 10.82, "Sasthamkotta": 11.17, "Mynagappally": 7.75,
        "Thazhava": 4.73, "Clappana": 8.41, "Vallikavu": 10.31, "Amritapuri": 11.60, "Puthiyakavu": 0.0
    }
}

# Origin trip generation propensity weights (P_i)
POPULATION_WEIGHTS = {
    "Karunagappally": 28.0,
    "Railway Station": 20.0,
    "KSRTC Bus Stand": 18.0,
    "Oachira": 14.0,
    "Chavara": 15.0,
    "Sasthamkotta": 12.0,
    "Mynagappally": 8.0,
    "Thazhava": 8.0,
    "Clappana": 6.0,
    "Vallikavu": 10.0,
    "Amritapuri": 18.0,
    "Puthiyakavu": 11.0,
}

# Destination trip attraction weights (A_j)
ATTRACTION_WEIGHTS = {
    "Karunagappally": 26.0,
    "Railway Station": 22.0,
    "KSRTC Bus Stand": 19.0,
    "Oachira": 13.0,
    "Chavara": 14.0,
    "Sasthamkotta": 11.0,
    "Mynagappally": 7.0,
    "Thazhava": 7.0,
    "Clappana": 5.0,
    "Vallikavu": 9.0,
    "Amritapuri": 20.0,
    "Puthiyakavu": 10.0,
}

# Hourly trip probability weights (24 hours: peak morning, evening, and off-peak night)
HOURLY_WEIGHTS = np.array([
    0.010, 0.008, 0.006, 0.008, 0.015, 0.025,  # 00:00 - 05:00 Early morning / Night
    0.045, 0.070, 0.085, 0.080, 0.055, 0.045,  # 06:00 - 11:00 Morning Peak
    0.040, 0.040, 0.045, 0.055, 0.075, 0.085,  # 12:00 - 17:00 Midday / Evening onset
    0.080, 0.065, 0.045, 0.030, 0.020, 0.013   # 18:00 - 23:00 Evening Peak & Night
], dtype=float)
HOURLY_WEIGHTS /= HOURLY_WEIGHTS.sum()


def build_road_graph():
    """Builds a simplified 12-landmark topological road network for graph-mining comparison."""
    graph = nx.Graph(kind="simulated_road_network")
    for name in LOCATIONS:
        graph.add_node(name, x=LAYOUT[name][0], y=LAYOUT[name][1])
    for s_idx, s in enumerate(LOCATIONS):
        for t in LOCATIONS[s_idx + 1:]:
            dist = EXACT_ROAD_DISTANCES[s][t]
            if dist < 12.0:  # Connect reasonably proximate landmark pairs
                time_min = round(dist * (60.0 / 28.0) + 2.0, 1)
                fare = round(FARE_BASE + dist * FARE_PER_KM + time_min * FARE_PER_MIN, 1)
                graph.add_edge(s, t, distance_km=float(dist), time_min=time_min, fare=float(fare))
    return graph


def compute_gravity_matrix(gamma=1.3):
    """Computes the 12x12 origin-destination probability matrix using the Gravity Model."""
    n = len(LOCATIONS)
    od_matrix = np.zeros((n, n), dtype=float)
    for i, orig in enumerate(LOCATIONS):
        for j, dest in enumerate(LOCATIONS):
            if i == j:
                od_matrix[i, j] = 0.0
            else:
                dist = max(0.5, EXACT_ROAD_DISTANCES[orig][dest])
                mass = POPULATION_WEIGHTS[orig] * ATTRACTION_WEIGHTS[dest]
                od_matrix[i, j] = mass / (dist ** gamma)
    total_mass = od_matrix.sum()
    return od_matrix / total_mass


def make_trips(n=TRIP_COUNT, seed=SEED, gamma=1.3):
    """Generates synthetic taxi trips deterministically using the gravity model."""
    rng = np.random.default_rng(seed)
    prob_matrix = compute_gravity_matrix(gamma=gamma)
    num_locs = len(LOCATIONS)

    # Sample origin-destination index pairs
    flat_indices = rng.choice(num_locs * num_locs, size=n, p=prob_matrix.flatten())
    orig_indices = flat_indices // num_locs
    dest_indices = flat_indices % num_locs

    origins = [LOCATIONS[i] for i in orig_indices]
    destinations = [LOCATIONS[j] for j in dest_indices]

    # Compute distances from exact road network matrix
    distances = np.array([EXACT_ROAD_DISTANCES[o][d] for o, d in zip(origins, destinations)], dtype=float)

    # Travel time estimation assuming 28 km/h typical peri-urban speed + 2 min boarding/intersection delay
    travel_times = np.round(distances * (60.0 / 28.0) + 2.0, 1)

    # Kerala MVD-grounded fare calculation: Base (45) + 18/km + 1.5/min
    fares = np.round(FARE_BASE + distances * FARE_PER_KM + travel_times * FARE_PER_MIN).astype(int)

    # Sample temporal timestamps over a 30-day window following diurnal hourly weights
    base_timestamp = pd.Timestamp("2026-09-01 00:00:00")
    days = rng.integers(0, 30, size=n)
    hours = rng.choice(24, size=n, p=HOURLY_WEIGHTS)
    minutes = rng.integers(0, 60, size=n)
    seconds = rng.integers(0, 60, size=n)

    dates = [
        base_timestamp + pd.Timedelta(days=int(d), hours=int(h), minutes=int(m), seconds=int(s))
        for d, h, m, s in zip(days, hours, minutes, seconds)
    ]

    return pd.DataFrame({
        "trip_id": [f"RTW-{i:05d}" for i in range(1, n + 1)],
        "pickup": origins,
        "dropoff": destinations,
        "datetime": dates,
        "distance_km": distances.round(2),
        "travel_time_min": travel_times,
        "estimated_fare_inr": fares
    })


def build_demand_graph(trips):
    """Builds a directed, weighted demand graph where edge weights represent trip frequencies."""
    graph = nx.DiGraph(kind="synthetic_taxi_demand")
    pickups = trips.pickup.value_counts()
    dropoffs = trips.dropoff.value_counts()
    for location in LOCATIONS:
        graph.add_node(
            location,
            pickups=int(pickups.get(location, 0)),
            dropoffs=int(dropoffs.get(location, 0)),
            x=LAYOUT[location][0],
            y=LAYOUT[location][1]
        )
    for (a, b), count in trips.groupby(["pickup", "dropoff"]).size().items():
        graph.add_edge(a, b, trips=int(count))
    return graph
