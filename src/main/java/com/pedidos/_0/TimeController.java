package com.pedidos._0;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.Map;

@RestController
public class TimeController {

    @GetMapping("/api/hora")
    public Map<String, String> getHora() {
        LocalDateTime ahora = LocalDateTime.now();

        // Formatear la hora en HH:mm:ss
        DateTimeFormatter formatoHora = DateTimeFormatter.ofPattern("HH:mm:ss");
        String horaString = ahora.format(formatoHora);

        // Formatear la fecha en dd/MM/yyyy
        DateTimeFormatter formatoFecha = DateTimeFormatter.ofPattern("dd/MM/yyyy");
        String fechaString = ahora.format(formatoFecha);

        // Construir el mensaje solicitado
        String mensaje = "El servidor marca las : " + horaString + " del " + fechaString;

        // Retornar la respuesta como un objeto Map que Spring Boot convierte
        // automáticamente a JSON
        Map<String, String> response = new HashMap<>();
        response.put("mensaje", mensaje);

        return response;
    }
}
