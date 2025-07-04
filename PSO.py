import random

def f(x):
    # Ejemplo: función parabólica
    return (x - 2)**2
# Parámetros del PSO

w = 0.7       # peso de inercia
c1 = 1.5      # coeficiente cognitivo (experiencia propia)
c2 = 1.5      # coeficiente social (experiencia global)
n_particles = 20
n_iter = 10
x_min, x_max = 0, 10
v_max = (x_max - x_min) * 0.1
v_min = -v_max


# Posiciones y velocidades iniciales
particles = [random.uniform(x_min, x_max) for _ in range(n_particles)]
velocities = [random.uniform(v_min, v_max) for _ in range(n_particles)]
pbest = particles.copy()
pbest_val = [f(x) for x in particles]
gbest = pbest[pbest_val.index(min(pbest_val))]
gbest_val = min(pbest_val)


for t in range(n_iter):
    for i in range(n_particles):
        r1, r2 = random.random(), random.random()
        velocities[i] = (
            w * velocities[i] +
            c1 * r1 * (pbest[i] - particles[i]) +
            c2 * r2 * (gbest - particles[i])
        )
        # Limitar velocidad
        velocities[i] = max(v_min, min(v_max, velocities[i]))
        
        # Actualizar posición con límites
        particles[i] += velocities[i]
        particles[i] = max(x_min, min(x_max, particles[i]))
        
        # Evaluar
        val = f(particles[i])
        if val < pbest_val[i]:  # para minimización
            pbest[i], pbest_val[i] = particles[i], val
            if val < gbest_val:
                gbest, gbest_val = particles[i], val

print(f"Mejor solución: x = {gbest}, f(x) = {gbest_val}")
