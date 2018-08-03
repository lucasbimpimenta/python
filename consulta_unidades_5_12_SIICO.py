# -*- coding: utf-8 -*-
from py3270 import Emulator

em = Emulator(visible=True)

# o codigo das unidades devem ter 4 digitos
unidades = ["0002", "5229", "5262", "5795"]

em.connect('L:ibmbr10.coredf.caixa:2301')
em.wait_for_field()

# Entrar no 5.12
em.send_string("512")
em.send_enter()
em.wait_for_field()

# Opcao 1 de consulta pelo codigo da unidade
em.send_string("1")
em.send_enter()
em.wait_for_field()

print("CGC;Tipo;Nome;Sigla;UF;Subordinacao")
# Consultando as unidades
for unid in unidades:
    em.send_string(unid)
    em.send_enter()
    em.wait_for_field()
    print(
        "{};{};{};{};{};{}"
        .format(
            unid,                               # CGC
            em.string_get(5, 45, 2),            # Tipo
            em.string_get(6, 18, 36).strip(),   # Nome
            em.string_get(6, 70, 8).strip(),    # Sigla
            em.string_get(10, 70, 2),           # UF
            em.string_get(14, 18, 4)            # Subordinacao
        )
    )

em.terminate()
