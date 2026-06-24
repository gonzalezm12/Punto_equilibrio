# 🤖 AI Rules: Plan de Auditoría Arquitectónica y Refactorización Backend-First

Este documento establece las **REGLAS ABSOLUTAS** de comportamiento y arquitectura para cualquier IA o desarrollador que trabaje en este proyecto ("Gestión Detrás del Balance"). 

Ninguna línea de código nueva o refactorización puede romper las siguientes directrices:

---

## 1. 🧠 La Regla de "UI Tonta Total" (Dumb UI)

La interfaz gráfica (React/Next.js) **no debe pensar, no debe consultar bases de datos directamente, ni hacer cálculos matemáticos**. Solo debe mapear el estado que le llega del Controller y despachar eventos de usuario hacia el Controller.

*   **Toda lógica, validación de negocio y cálculos** debe estar exclusivamente en **Servicios de Dominio** puros (`src/modulos/<dominio>/dominio/reglas.ts`).
*   Los **Use Cases (Casos de Uso)** solo orquestan. Su única misión es inyectar un Repositorio, invocar la validación del Dominio y guardar el resultado. No contienen algoritmos financieros.
*   Los **Hooks (`useController`)** deben ser puros. Actúan como "pasamanos". Toman datos del estado global o formularios y llaman a un UseCase. No pueden sumar saldos ni validar reglas.
*   **Provider, Controller y Context** solo manejan estado (React Context / UI State).
*   **Llamadas a Firebase (`doc`, `getDoc`, `collection`)** están estrictamente **prohibidas** en archivos `.tsx` de UI. Solo los Repositorios en `infraestructura/` pueden importar de `firebase/firestore`.
*   **Indicar Infracciones:** Si el usuario solicita un cambio que rompe estas reglas, la IA debe indicar exactamente dónde se rompe la regla y por qué, y proponer la ruta limpia.

---

## 2. ☁️ Arquitectura Backend-First (Alta Escalabilidad y Concurrencia)

Para evitar la doble contabilización (double-counting), proteger los datos contra race conditions, evitar costos masivos de lectura y soportar iteraciones de múltiples usuarios escribiendo al milisegundo:

*   **Cálculos Pesados y Agregaciones en el Backend:** La suma de `monthlyTotals` (Totales mensuales), saldos de Cuentas, balances de Presupuestos y resultados operativos se calculan de manera atómica mediante **Cloud Functions** (`transactionTriggers.js` a través del evento `onWrite`).
*   **Transacciones Atómicas (UoW):** Cualquier impacto en el Ledger o múltiples colecciones debe envolverse en el patrón Unit of Work.
*   **Lecturas Óptimas:** El Frontend no debe descargar ni iterar sobre cientos de documentos para obtener un saldo total. Se debe leer el pre-agregado directamente.

---

## 3. 🛡️ Principios SOLID Inquebrantables

*   **SRP (Responsabilidad Única):** Está prohibida la creación de "God Classes" (Servicios que hagan persistencia, notificaciones y cálculos al mismo tiempo).
*   **OCP (Abierto/Cerrado):** Reemplazar cascadas de `if/else` o `switch` por Polimorfismo y Patrón Strategy cuando se procesen tipos distintos (ej. Ingresos vs Egresos).
*   **DIP (Inversión de Dependencias):** El código de negocio (`aplicacion/`) jamás importará la tecnología concreta. Siempre debe depender de abstracciones (ej. inyectar `ITransactionRepository`).
*   **ISP (Segregación de Interfaces):** Los Contextos globales no deben exponer mutaciones masivas si el componente solo requiere lectura.

---

## 4. 📏 Límites Estrictos de Código

*   **Ningún archivo puede superar las 250 - 300 líneas de código como máximo.** Si un archivo rebasa este límite, es una señal obligatoria de que rompe el SRP y debe ser fragmentado obligatoriamente "siempre dividir, sin introducir regresiones".

---

## 5. 🛑 Política de Cero Regresiones

Toda refactorización debe ejecutarse garantizando que **no se introduzcan regresresiones en la UI, ni en los cálculos, validaciones o reglas de negocio**.

1. Extraer la lógica pura a funciones independientes (Dominio).
2. Dejar el servicio/componente original como una *Fachada (Facade)* si es necesario para estabilizar las firmas de las funciones que usa React.
3. Asegurar mediante ejecución de TypeScript (`tsc`) que el código sea estructuralmente sano y sin errores.
4. Jamás introducir "cabos sueltos". El código muerto debe ser purgado.

---

# 🔥 NUEVAS REGLAS AGREGADAS

---

## 6. 🧩 Agnosticismo Tecnológico (Independencia del Backend)

La aplicación **no debe depender de Firebase ni de ninguna tecnología específica**.  
Toda la lógica debe ser portable a cualquier proveedor de base de datos o backend.

* Los **Repositorios** deben ser interfaces (`IUserRepository`, `ITransactionRepository`, etc.) ubicadas en `src/aplicacion/repositorios/`.
* La capa de **Dominio** no puede importar Firebase, Firestore, Supabase, SQL, ni ninguna tecnología concreta.
* La capa de **Infraestructura** implementa los repositorios concretos (FirebaseRepository, SupabaseRepository, SQLRepository).
* La UI y los Casos de Uso **solo dependen de interfaces**, nunca de implementaciones.
* La IA debe advertir si se introduce cualquier dependencia directa a Firebase fuera de `infraestructura/`.

---

## 7. 🔐 Regla de Seguridad: Cero Credenciales en el Código

Está estrictamente prohibido:

* Hardcodear claves API, tokens, secrets, JSON de Firebase, credenciales o rutas privadas.
* Incluir archivos `.json` de Firebase Admin en el repositorio.
* Exponer datos de usuarios, correos, IDs o información sensible en el código.
* Incluir tokens en ejemplos de código generados por la IA.

Todo secreto debe estar en:

* `.env.local` (desarrollo)  
* Variables de entorno del servidor (producción)  
* Secret Manager (Firebase / Cloudflare / AWS)

La IA debe advertir si el usuario solicita algo que viole esta regla.

---

## 8. 🏗️ Regla de Estructura Escalable del Proyecto

La estructura del proyecto debe ser **modular, escalable y preparada para crecimiento futuro**, incluyendo apps móviles y microservicios.

Estructura mínima obligatoria:

