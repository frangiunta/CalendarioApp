# --- ETAPA 1: Construcción (Builder) ---
FROM node:20-alpine AS builder

WORKDIR /app

# Copiamos todo el código fuente al contenedor
COPY . .

# Eliminamos el package-lock.json por si está corrupto y reinstalamos dependencias
RUN rm -f package-lock.json && npm install

# Compilamos la aplicación de Angular para producción
RUN npm run build

# --- ETAPA 2: Servidor Web ---
FROM nginx:alpine

# Limpiar la carpeta por defecto de Nginx
RUN rm -rf /usr/share/nginx/html/*

# Copiar el paquete ya listo desde la Etapa 1
COPY --from=builder /app/dist/app-francos/browser /usr/share/nginx/html

EXPOSE 80
CMD ["nginx", "-g", "daemon off;"]