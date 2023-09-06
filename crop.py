from pdfplumber.page import Page

def crop_several(page: Page, i: int) -> Page:
    match i:
        case 0:
            return crop_seguro(page)
        
        case 1:
            return crop_dep(page)

        case 2:
            return crop_nome(page)

        case 3:
            return crop_regfunc(page)
        
        case 4:
            return crop_idade(page)
        
        case 5:
            return crop_parentesco(page)
        
        case 6:
            return crop_plano(page)
        
        case 7:
            return crop_mov(page)
        
        case 8:
            return crop_situacao(page)

def crop_left_side(page: Page) -> Page:
    return page.within_bbox((0, 130, 690, 595))  # X0, Y0, X1, Y1

def crop_seguro(page: Page) -> Page:
    return page.within_bbox((0, 130, 70, 595))

def crop_dep(page: Page) -> Page:
    return page.within_bbox((60, 130, 100, 595))

def crop_nome(page: Page) -> Page:
    return page.within_bbox((100, 130, 250, 595))

def crop_regfunc(page: Page) -> Page:
    return page.within_bbox((250, 130, 300, 595))

def crop_idade(page: Page) -> Page:
    return page.within_bbox((300, 130, 330, 595))

def crop_parentesco(page: Page) -> Page:
    return page.within_bbox((330, 130, 380, 595))

def crop_plano(page: Page) -> Page:
    return page.within_bbox((380, 130, 540, 595))

def crop_mov(page: Page) -> Page:
    return page.within_bbox((540, 130, 610, 595))

def crop_situacao(page: Page) -> Page:
    return page.within_bbox((610, 130, 690, 595))

def crop_right_side(page: Page) -> Page:
    return page.within_bbox((690, 130, 842, 595))

def crop_headers(page: Page) -> Page:
    return page.within_bbox((0, 113, 842, 123))

def crop_qe_valores(page: Page) -> Page:
    return page.within_bbox((100, 138, 415, 215))

def crop_premio(page: Page) -> Page:
    return page.within_bbox((600, 138, 830, 215))

def crop_totalizador(page: Page) -> Page:
    return page.within_bbox((445, 238, 830, 345))

def crop_detalhamento(page: Page) -> Page:
    return page.within_bbox((0, 380, 842, 595))
