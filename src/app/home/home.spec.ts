import { ComponentFixture, TestBed } from '@angular/core/testing';

import { homeComponent } from './home';

describe('Home', () => {
  let component: homeComponent;
  let fixture: ComponentFixture<homeComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [homeComponent],
    }).compileComponents();

    fixture = TestBed.createComponent(homeComponent);
    component = fixture.componentInstance;
    await fixture.whenStable();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
