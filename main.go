package main

import (
	"fmt"
	"time"
)

func compter(nom string) {
	for i := 1; i <= 3; i++ {
		fmt.Printf("🔄 %s : étape %d\n", nom, i)
		time.Sleep(1 * time.Second)
	}
}

func main() {
	fmt.Println("🚀 Démarrage du test de concurrence...")

	// On lance la fonction en arrière-plan
	go compter("Dionysos-Background") 

	// On lance la fonction au premier plan
	compter("Digixtp-Main")

	fmt.Println("✅ Test terminé !")
}