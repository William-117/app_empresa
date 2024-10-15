
from app import app
from flask import render_template, request, flash, redirect, url_for, session

# importando mi conexion a BD
from config import connectionBD

# para encriptar contraseña generate_password_hash
from werkzeug.security import check_password_hash

# importando controllers para el modulo de login 
from controllers.funciones_usuario import *
PATH_URL_LOGIN= "login" # carpeta en templates/login


@app.route('/', methods=['GET'])
def inicio():
    if 'conectado' in session:
        return render_template('base.html', dataLogin=dataLoginSesion())
    else:
        return render_template(f'{PATH_URL_LOGIN}/base_login.html')


@app.route('/mi-perfil', methods=['GET'])
def perfil():
    if 'conectado' in session:
        return render_template(f'perfil/perfil.html', info_perfil_session=info_perfil_session())
    else:
        return redirect(url_for('inicio'))


# crear cuenta de usuario
@app.route('/register-user', methods=['GET'])
def cpanelRegisterUser():
    if 'conectado' in session:
        return redirect(url_for('inicio'))
    else:
        return render_template(f'{PATH_URL_LOGIN}/auth_register.html')
    

#recuperar cuenta de usuario
@app.route('/recovery-password', methods=['GET'])
def cpanelRecoveryPassUser():
    if 'conectado' in session:
        return redirect(url_for('inicio'))
    else:
        return render_template(f'{PATH_URL_LOGIN}/auth_forgot_password.html')


#crear cuenta de usuario
@app.route('/saved-register', methods=['POST'])
def cpanelResgisterUserBD():
    if request.method == 'POST' and 'usuario' in request.form and 'password' in request.form:
        usuario = request.form['usuario']
        email = request.form['email']
        password = request.form['password']

        resultData = recibeInsertRegisterUser(
            usuario, email, password)
        if (resultData != 0):
            flash('la cuenta fue creada correctamente.', 'success')
            return redirect(url_for('inicio'))
        else:
            return redirect(url_for('inicio'))
    else:
        flash('el método HTTP es incorrecto', 'error')
        return redirect(url_for('inicio'))
    

#actualizar datos de mi perfil
@app.route('/actualizar-datos-perfil', methods=['POST'])
def actualizarPerfil():
    if request.method == 'POST':
        if 'conectado' in session:
            respuesta = procesar_update_perfil(request.form)
            if respuesta == 1:
                flash('Los datos fueron actualizados correctamente.', 'success')
                return redirect(url_for('inicio'))
            elif respuesta == 0:
                flash(
                    'La contraseña actual esta incorrecta, por favor verifique.', 'error')
                return redirect(url_for('perfil'))
            elif respuesta == 2:
                flash('Ambas claves deben ser igual, por favor verifique.', 'error')
                return redirect(url_for('perfil'))
            elif respuesta == 3:
                flash('La Clave actual es obligatoria.', 'error')
                return redirect(url_for('perfil'))
        else:
            flash('primero debes iniciar sesión.', 'error')
            return redirect(url_for('inicio'))
    else:
        flash('primero debes iniciar sesión.', 'error')
        return redirect(url_for('inicio'))


#validar sesion
@app.route('/login', methods=['GET', 'POST'])
def loginCliente():
    if 'conectado' in session:
        return redirect(url_for('inicio'))
    else:
        if request.method == 'POST' and 'email' in request.form and 'password' in request.form:
            
            email = str(request.form['email'])
            password = str(request.form['password'])

            # Comprobando si existe una cuenta
            conexion_MySQLdb = connectionBD()
            cursor = conexion_MySQLdb.cursor(dictionary=True)
            cursor.execute(
                "SELECT * FROM usuarios WHERE email = %s", [email])
            account = cursor.fetchone()

            if account:
                if check_password_hash(account['password'], password):
                    # Crear datos de sesión, para poder acceder a estos datos en otras rutas
                    session['conectado'] = True
                    session['id_usuario'] = account['id_usuario']
                    session['usuario'] = account['usuario']
                    session['email'] = account['email']

                    flash('la sesión fue correcta.', 'success')
                    return redirect(url_for('inicio'))
                else:
                    # La cuenta no existe o el nombre de usuario/contraseña es incorrecto
                    flash('datos incorrectos por favor revise.', 'error')
                    return render_template(f'{PATH_URL_LOGIN}/base_login.html')
            else:
                flash('el usuario no existe, por favor verifique.', 'error')
                return render_template(f'{PATH_URL_LOGIN}/base_login.html')
        else:
            flash('primero debes iniciar sesión.', 'error')
            return render_template(f'{PATH_URL_LOGIN}/base_login.html')
        

@app.route('/closed-session', methods=['GET'])
def cerraSesion():
    if request.method == 'GET':
        if 'conectado' in session:
            # Eliminar datos de sesión, esto cerrará la sesión del usuario
            session.pop('conectado', None)
            session.pop('id_usuario', None)
            session.pop('usuario', None)
            session.pop('email', None)
            flash('tu sesión fue cerrada correctamente.', 'success')
            return redirect(url_for('inicio'))
        else:
            flash('recuerde debe iniciar sesión.', 'error')
            return render_template(f'{PATH_URL_LOGIN}/base_login.html')