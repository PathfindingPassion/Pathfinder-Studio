# Pathfinder-Studio
Pathfinder Studio is a high performance pathfinding engine collection
# Pathfinder Studio v3.1 🧭

A fast, interactive, and customizable 2D/3D pathfinding algorithm visualizer and solver built entirely in **Microsoft Excel (VBA)**.

---

## 📌 Overview

**Pathfinder Studio** is designed to demonstrate, analyze, and solve network pathfinding problems directly inside Excel. By utilizing optimized VBA implementations of classic graph algorithms, it allows users to model grids, calculate shortest paths, and visualize node expansions in real time.

Whether you are working on grid-based routing, spatial optimization, or educational demonstrations, Pathfinder Studio provides a seamless interface within a spreadsheet environment.

---

## ✨ Key Features

* **Dual Core Algorithms:**
  * **Dijkstra's Algorithm:** Standard, robust shortest-path search for general weighted graphs.
  * **Dial's Algorithm (Radix Heap optimization):** High-performance pathfinding optimized for small integer edge weights.
* **2D & 3D Grid Support:** Solve pathfinding problems across traditional 2D maps or multi-layered 3D structures.
* **Interactive Visualization:** Observe propagation steps directly on the Excel worksheet grid.
* **Obstacle & Weight Customization:** Define impassable barriers, custom movement costs, and diagonal paths.
* **No Dependencies:** Runs natively in Microsoft Excel with VBA enabled—no external libraries or plugins required.

---

## 🛠️ How to Use

1. **Download the Workbook:**
   * Download the `.xlsm` (or `.xlsb`) file from this repository.
2. **Enable Macros:**
   * Open the file in Excel and ensure VBA Macros are enabled (**Enable Content**).
3. **Set Up Your Grid:**
   * Define your **Start** and **End** nodes on the worksheet grid.
   * Add walls/obstacles or custom movement costs as needed.
4. **Run the Solver:**
   * Select your preferred algorithm (Dijkstra or Dial) and execution settings.
   * Click **Start / Solve** to launch the pathfinding process and display the shortest path.

---

## 📊 Performance & Optimization

Pathfinder Studio has been optimized in native VBA using structured array processing and efficient data structures to minimize worksheet rendering latency and deliver high computation speeds even on large grids.

---

## 👨‍💻 Author & Contact

* **Author:** Stéphane Fillon
* **Feedback & Contributions:** Feel free to open an **Issue** or submit a **Pull Request** if you have suggestions, bug reports, or feature requests!

---

*If you find this project useful or interesting, please consider giving it a ⭐ on GitHub!*
