/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log("Office.js chargé, add-in prêt");
  }
});

/**
 * Recherche LinkedIn avec le nom de l'expéditeur ou du contact
 * Cette fonction est appelée depuis le manifest via ExecuteFunction
 * IMPORTANT: Cette fonction doit être dans le scope global pour être accessible
 */
function searchLinkedIn(event) {
  const item = Office.context.mailbox.item;
  
  try {
    let fullName = null;
    
    // Vérifier le type d'élément
    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
      // Pour un message (mail), récupérer l'expéditeur
      const sender = item.from;
      if (sender) {
        fullName = sender.displayName || sender.emailAddress || null;
      }
    } 
    else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
      // Pour un rendez-vous, récupérer l'organisateur
      const organizer = item.organizer;
      if (organizer) {
        fullName = organizer.displayName || organizer.emailAddress || null;
      }
      
      // Si pas d'organisateur, essayer les participants requis
      if (!fullName) {
        const requiredAttendees = item.requiredAttendees;
        if (requiredAttendees && requiredAttendees.length > 0) {
          fullName = requiredAttendees[0].displayName || requiredAttendees[0].emailAddress || null;
        }
      }
      
      // Si toujours pas de nom, essayer les participants optionnels
      if (!fullName) {
        const optionalAttendees = item.optionalAttendees;
        if (optionalAttendees && optionalAttendees.length > 0) {
          fullName = optionalAttendees[0].displayName || optionalAttendees[0].emailAddress || null;
        }
      }
    }
    
    if (fullName) {
      // Nettoyer le nom (enlever les adresses email si c'est juste un email)
      if (fullName.includes('@') && !fullName.includes(' ')) {
        fullName = fullName.split('@')[0].replace(/[._]/g, ' ');
      }
      
      // Encoder le nom pour l'URL
      const encodedName = encodeURIComponent(fullName.trim());
      // Construire l'URL de recherche LinkedIn
      const linkedInUrl = `https://www.linkedin.com/search/results/people/?keywords=${encodedName}`;
      
      // Ouvrir LinkedIn dans le navigateur par défaut du système
      // Office.context.ui.openBrowserWindow() (Mailbox 1.6+) ouvre dans le navigateur par défaut
      openInDefaultBrowser(linkedInUrl, event);
    } else {
      showNotification("Impossible de récupérer le nom. Veuillez sélectionner un mail ou un contact.");
      event.completed();
    }
  } catch (error) {
    console.error("Erreur lors de la recherche LinkedIn:", error);
    showNotification("Une erreur est survenue: " + error.message);
    event.completed();
  }
}

/**
 * Ouvre une URL dans le navigateur par défaut du système
 * Utilise openBrowserWindow (Mailbox 1.6+) avec fallback
 */
function openInDefaultBrowser(url, event) {
  // Vérifier si openBrowserWindow est disponible (Mailbox 1.6+)
  if (Office.context.ui.openBrowserWindow) {
    Office.context.ui.openBrowserWindow(url);
    console.log("LinkedIn ouvert dans le navigateur par défaut:", url);
    event.completed();
  } else {
    // Fallback pour les anciennes versions: utiliser displayDialogAsync avec redirection
    console.warn("openBrowserWindow non disponible, utilisation du fallback");
    
    // Créer une page de redirection dynamique
    const redirectHtml = `https://rise-4.github.io/all--outlook-linkedin--addin/redirect.html?url=${encodeURIComponent(url)}`;
    
    Office.context.ui.displayDialogAsync(
      redirectHtml,
      { height: 10, width: 10, displayInIframe: false },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          const dialog = asyncResult.value;
          // Fermer le dialogue après un court délai
          setTimeout(() => {
            try {
              dialog.close();
            } catch (e) {
              // Dialogue peut déjà être fermé
            }
          }, 2000);
        } else {
          console.error("Erreur displayDialogAsync:", asyncResult.error);
          showNotification("Impossible d'ouvrir LinkedIn. URL: " + url);
        }
        event.completed();
      }
    );
  }
}

/**
 * Affiche une notification à l'utilisateur
 * Utilise l'API de notification d'Outlook si disponible
 */
function showNotification(message) {
  // Utiliser l'API de notification si disponible (plus propre que displayDialogAsync)
  if (Office.context.mailbox.item && Office.context.mailbox.item.notificationMessages) {
    Office.context.mailbox.item.notificationMessages.replaceAsync(
      "linkedin-notification",
      {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: message,
        icon: "Icon.16x16",
        persistent: false
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error("Erreur notification:", result.error);
          // Fallback: alert simple
          console.log("Message pour l'utilisateur:", message);
        }
      }
    );
  } else {
    console.log("Message pour l'utilisateur:", message);
  }
}

