import argparse
import sys
from presentation import creer, openFile


def main():
    parser = argparse.ArgumentParser(
        description="Créer un fichier PPTX à partir d'arguments et l'ouvrir"
    )
    parser.add_argument("id", type=str, help="ID")
    parser.add_argument("toko_tena", type=str, help="toko_tena")
    parser.add_argument("andininydeb_tena", type=str, help="andininydeb_tena")
    parser.add_argument("andininyfin_tena", type=str, help="andininyfin_tena")

    try:
        args = parser.parse_args()
        fichier = creer(
            args.id, args.toko_tena, args.andininydeb_tena, args.andininyfin_tena
        )
        openFile(fichier)
    except Exception as e:
        print(str(e))
        sys.exit(1)


if __name__ == "__main__":
    main()
