import { Component, OnInit, ChangeDetectorRef } from '@angular/core';
import { RouterOutlet, Router } from '@angular/router';
import { CommonModule, Location } from '@angular/common';
import { MsalBroadcastService, MsalService, MsalCustomNavigationClient } from '@azure/msal-angular';
import { InteractionRequiredAuthError, InteractionStatus } from '@azure/msal-browser';
import { HttpClient } from '@angular/common/http';
import { filter } from 'rxjs';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [RouterOutlet, CommonModule],
  styleUrl: './app.css',
  templateUrl: './app.html',
})
export class App implements OnInit {
  title = 'Pedidos360';
  isLoggedIn = false;
  userName = '';
  apiResponse = '';
  apiError = '';

  constructor(
    private authService: MsalService,
    private msalBroadcastService: MsalBroadcastService,
    private http: HttpClient,
    private router: Router,
    private location: Location,
    private cdr: ChangeDetectorRef
  ) { }

  async ngOnInit(): Promise<void> {
    try {
      // 1. Inicializar MSAL
      await this.authService.instance.initialize();

      // Configurar NavigationClient personalizado para que Angular maneje la navegación interna
      // Esto evita que MSAL ejecute window.location.replace() recargando toda la ventana y causando loops
      const customNavigationClient = new MsalCustomNavigationClient(
        this.authService,
        this.router,
        this.location
      );
      this.authService.instance.setNavigationClient(customNavigationClient);

      // 2. Procesar el retorno de la redirección desde Microsoft Entra ID
      this.authService.handleRedirectObservable({ navigateToLoginRequestUrl: false }).subscribe({
        next: (result) => {
          console.log('MSAL Redirect resultado:', result);
          if (result?.account) {
            this.authService.instance.setActiveAccount(result.account);
          }
          this.verificarCuenta();
        },
        error: (error) => {
          console.error('Error procesando autenticación:', error);
        }
      });

      // 3. Escuchar cuando MSAL termine de procesar todos sus estados internos
      this.msalBroadcastService.inProgress$
        .pipe(filter((status: InteractionStatus) => status === InteractionStatus.None))
        .subscribe(() => {
          this.verificarCuenta();
        });

      // 4. Verificación inicial por si ya había sesión abierta previamente
      this.verificarCuenta();

    } catch (error) {
      console.error('Error inicializando MSAL:', error);
    }
  }

  verificarCuenta(): void {
    let activeAccount = this.authService.instance.getActiveAccount();

    if (!activeAccount) {
      const accounts = this.authService.instance.getAllAccounts();
      if (accounts.length > 0) {
        activeAccount = accounts[0];
        this.authService.instance.setActiveAccount(activeAccount);
      }
    }

    if (activeAccount) {
      this.isLoggedIn = true;
      this.userName = activeAccount.name || activeAccount.username;
    } else {
      this.isLoggedIn = false;
      this.userName = '';
    }
    this.cdr.detectChanges();
  }

  login(): void {
    this.authService.loginRedirect();
  }

  logout(): void {
    this.authService.logoutRedirect({
      postLogoutRedirectUri: 'http://localhost:4200/'
    });
  }

  async consultarHora(): Promise<void> {
    this.apiResponse = '';
    this.apiError = '';

    const apiUrl =
      'https://f13t0uil1i.execute-api.us-east-1.amazonaws.com/api/hora';

    const account =
      this.authService.instance.getActiveAccount();

    if (!account) {
      this.apiError = 'No existe una sesión autenticada.';
      return;
    }

    const tokenRequest = {
      scopes: [
        'api://ab7113ad-620d-4f62-8dc5-9da78f45bf8d/access_as_user'
      ],
      account: account
    };

    try {
      console.log('Solicitando access token para Pedidos360...');

      const tokenResponse =
        await this.authService.instance.acquireTokenSilent(
          tokenRequest
        );

      console.log('Access token obtenido correctamente.');

      const partes = tokenResponse.accessToken.split('.');
      const payload = JSON.parse(atob(partes[1]));

      console.log('ISSUER DEL TOKEN:', payload.iss);
      console.log('AUDIENCE DEL TOKEN:', payload.aud);
      console.log('SCOPES DEL TOKEN:', payload.scp);


      const headers = {
        Authorization: `Bearer ${tokenResponse.accessToken}`
      };

      this.http.get<any>(
        apiUrl,
        { headers }
      ).subscribe({
        next: (response) => {
          console.log('Respuesta API:', response);

          if (response?.mensaje) {
            this.apiResponse = response.mensaje;
          } else {
            this.apiResponse = JSON.stringify(response);
          }
          this.cdr.detectChanges();
        },

        error: (err) => {
          console.error('Error consultando la API:', err);

          this.apiError =
            'Error al comunicarse con la API de AWS. ' +
            'Revisa la consola para más detalles.';
          this.cdr.detectChanges();
        }
      });

    } catch (error) {

      console.error(
        'Error obteniendo access token:',
        error
      );

      if (error instanceof InteractionRequiredAuthError) {
        this.apiError =
          'Microsoft Entra ID requiere autenticación o consentimiento adicional.';
      } else {
        this.apiError =
          'No fue posible obtener el access token para Pedidos360.';
      }
      this.cdr.detectChanges();
    }
  }

}