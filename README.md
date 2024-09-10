# happy-birthday
"""
Proyecto de calendario para cumpleaños hecho por Emilio Reato el 10 de septiembre de 2024.
Despliega un menu en la consola con el cual se puede operar y modificar los recordatorios.
Cuando llegue el día de uno de los cumpleaños registrados, sonará una notificacion de windows para que no se te olvide!.

He recibido la ayuda de chatgpt para investigar cuales librerias podrian funcionarme mejor y para crear ciertas funciones simples como check_registry_key_exists o ocultar_ventana.
El resultado final óptimo del proyecto es un .exe empaquetado gracias a auto_py_to_exe o pyinstaller --onedir que cuenta con una carpeta auxiliar(_internal) donde se guardan todos los archivos necesarios, aunque tambien funciona si se ejecuta el archivo .py (no se si en el startup)
"""
