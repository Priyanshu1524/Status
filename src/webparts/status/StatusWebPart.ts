import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

const arrowImage: string = require('./arrow.jpg');
import styles from './StatusWebPart.module.scss';

export interface ITicketStatusWebPartProps {
  description: string;
}

export default class TicketStatusWebPart extends BaseClientSideWebPart<ITicketStatusWebPartProps> {
  private _currentStatus: string = 'a';
  private _statusSteps: NodeListOf<Element>;

  public render(): void {
    const arr = ['Ticket Raised', 'a', 'b', 'c', 'd', 'e', 'Approval', 'Department Approval', 'IT', 'WIP', 'Closed'];
    let arrowHtml = `
      <div class="${styles.arrowWrapper}">
        <img src="${arrowImage}" alt="Arrow" class="${styles.arrowImage}">
      </div>
    `;
    const htmlstatus = arr.map((key, index) => {
      return `
        
          <div class="${styles.statusStep}">
            <div class="${styles.statusCircle}"></div>
            <div class="${styles.statusText}">${key}</div>
          </div>
        
      `;
    }).join(arrowHtml);
    
    this.domElement.innerHTML = `
      <div class="${styles.webpartContainer}">
        <div class="${styles.ticketContainer}">
          <div class="${styles.ticketNumber}">
            <label for="ticket-select" class="${styles.ticketNumberLabel}">Ticket No:</label>
            <select id="ticket-select" class="${styles.ticketSelect}">
              <option value="12">12</option>
              <option value="13">13</option>
              <option value="14">14</option>
            </select>
          </div>
          <div class="${styles.statusContainer}">
            ${htmlstatus}
          </div>
        </div>
      </div>
    `;

    this._statusSteps = this.domElement.querySelectorAll(`.${styles.statusStep}`);
    this.updateStatus(this._currentStatus);
  }

  protected onInit(): Promise<void> {
    return super.onInit().then(_ => {
      // Other initialization code can go here.
    });
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    // Update theme logic here if needed

    this.render();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "Ticket Status Web Part Configuration" },
          groups: [
            {
              groupName: "Settings",
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "Description"
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private updateStatus(currentStatus: string): void {
    let statusReached: boolean = false;
    this._statusSteps.forEach((step: Element) => {
      const statusText: string = step.querySelector(`.${styles.statusText}`)?.textContent || '';

      if (!statusReached) {
        step.classList.add(styles.statusStepActive);
      } else {
        step.classList.remove(styles.statusStepActive);
      }

      if (statusText === currentStatus) {
        statusReached = true;
        this.scrollToActiveStatus(step);
      }
    });
  }

  private scrollToActiveStatus(activeStep: Element): void {
    const container = this.domElement.querySelector(`.${styles.ticketContainer}`);
    if (container) {
      const activeStepRect = activeStep.getBoundingClientRect();
      const containerRect = container.getBoundingClientRect();
      const offset = activeStepRect.left - containerRect.left - (containerRect.width / 2) + (activeStepRect.width / 2);
      container.scrollLeft += offset;
    }
  }
}
